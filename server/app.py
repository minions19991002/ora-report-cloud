from __future__ import annotations

import os
import subprocess
import shutil
import sys
import tempfile
import threading
from calendar import monthrange
from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from fastapi import BackgroundTasks, FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from starlette.datastructures import UploadFile


APP_ROOT = Path(__file__).resolve().parent
STATIC_ROOT = APP_ROOT.parent / "static"
ENGINE_PATH = APP_ROOT / "engine" / "build_ora_report.py"

REQUIRED_FILES = {
    "previous": "ORA外送周报_6.8-6.14.xlsx",
    "template": "ora外送周报模版.xlsx",
    "storeInfo": "ORA门店信息表.xlsx",
    "mtStore": "美团门店数据.xlsx",
    "eleStore": "饿了么门店数据.xlsx",
    "mtPromo": "美团推广.xlsx",
    "elePromo": "饿了么推广.xlsx",
    "mtOrder": "美团订单数据.xlsx",
    "eleOrder": "饿了么订单数据.xlsx",
    "distance": "ora_订单距离分布_2026-06-22.xlsx",
    "mtProduct": "美团商品数据.xlsx",
    "eleProduct": "饿了么商品数据.xlsx",
    "reviewSummary": "ora_评价汇总_2026-06-15~2026-06-21.xlsx",
    "reviewCounts": "好评数中差评数据.xlsx",
    "delivery": "美团平均配送时长ora_自定义报表_2026-06-15_2026-06-21.xlsx",
    "oraDaily": "Ora外送日报.xlsx",
    "oraProduct": "Ora外送商品数据.xlsx",
}

ENV_KEYS = [
    "ORA_BASE",
    "ORA_WORK",
    "ORA_START",
    "ORA_END",
    "ORA_PREV_START",
    "ORA_PREV_END",
    "ORA_CURRENT_SHEET",
    "ORA_PREVIOUS_SHEET",
    "ORA_OUTPUT_NAME",
]

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GENERATE_LOCK = threading.Lock()
JOB_LOCK = threading.Lock()
JOB_ROOT = Path(tempfile.gettempdir()) / "ora-report-cloud-jobs"
JOBS: dict[str, dict[str, Any]] = {}
UPLOAD_SESSION_ROOT = Path(tempfile.gettempdir()) / "ora-report-cloud-upload-sessions"
UPLOAD_SESSION_LOCK = threading.Lock()
UPLOAD_SESSIONS: dict[str, dict[str, Any]] = {}
APP_VERSION = os.environ.get("RENDER_GIT_COMMIT") or os.environ.get("COMMIT_SHA") or "local"

app = FastAPI(title="ORA 外送报表生成服务")
app.mount("/assets", StaticFiles(directory=STATIC_ROOT), name="assets")


def period_label(start: str, end: str) -> str:
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)
    return f"{start_dt.month}.{start_dt.day}-{end_dt.month}.{end_dt.day}"


def export_period_label(start: str, end: str) -> str:
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)
    is_full_month_range = (
        start_dt.day == 1
        and end_dt.day == monthrange(end_dt.year, end_dt.month)[1]
        and end_dt >= start_dt
    )
    if is_full_month_range:
        if start_dt.year == end_dt.year and start_dt.month == end_dt.month:
            return f"{start_dt.month}月"
        if start_dt.year == end_dt.year:
            return f"{start_dt.month}-{end_dt.month}月"
        return f"{start_dt.year}.{start_dt.month}-{end_dt.year}.{end_dt.month}月"
    return period_label(start, end)


def report_kind(start: str, end: str) -> str:
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)
    is_full_month_range = (
        start_dt.day == 1
        and end_dt.day == monthrange(end_dt.year, end_dt.month)[1]
        and end_dt >= start_dt
    )
    return "月报" if is_full_month_range else "周报"


def require_date(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise HTTPException(status_code=400, detail=f"缺少周期字段：{field}")
    try:
        datetime.fromisoformat(value)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"日期格式不正确：{field}") from exc
    return value


def restore_env(previous: dict[str, str | None]) -> None:
    for key, value in previous.items():
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value


async def save_upload(upload: UploadFile, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("wb") as handle:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            handle.write(chunk)
    await upload.close()


def set_job(job_id: str, **updates: Any) -> None:
    with JOB_LOCK:
        current = JOBS.setdefault(job_id, {})
        current.update(updates)


def get_job(job_id: str) -> dict[str, Any]:
    with JOB_LOCK:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="任务不存在")
        return dict(job)


def get_upload_session(upload_id: str) -> dict[str, Any]:
    with UPLOAD_SESSION_LOCK:
        session = UPLOAD_SESSIONS.get(upload_id)
        if not session:
            raise HTTPException(status_code=404, detail="upload session not found")
        return dict(session)


def set_upload_session(upload_id: str, **updates: Any) -> None:
    with UPLOAD_SESSION_LOCK:
        current = UPLOAD_SESSIONS.setdefault(upload_id, {})
        current.update(updates)


def build_env(input_dir: Path, work_dir: Path, current_start: str, current_end: str, previous_start: str, previous_end: str) -> tuple[dict[str, str], str]:
    current_sheet = period_label(current_start, current_end)
    previous_sheet = period_label(previous_start, previous_end)
    output_name = f"ORA外送{report_kind(current_start, current_end)}_{export_period_label(current_start, current_end)}.xlsx"
    env = {
        "ORA_BASE": str(input_dir),
        "ORA_WORK": str(work_dir),
        "ORA_START": current_start,
        "ORA_END": current_end,
        "ORA_PREV_START": previous_start,
        "ORA_PREV_END": previous_end,
        "ORA_CURRENT_SHEET": current_sheet,
        "ORA_PREVIOUS_SHEET": previous_sheet,
        "ORA_OUTPUT_NAME": output_name,
    }
    return env, output_name


def run_engine(env: dict[str, str], work_dir: Path, output_name: str) -> Path:
    child_env = os.environ.copy()
    child_env.update(env)
    child_env["PYTHONIOENCODING"] = "utf-8"
    with GENERATE_LOCK:
        result = subprocess.run(
            [sys.executable, str(ENGINE_PATH)],
            cwd=str(APP_ROOT.parent),
            env=child_env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=900,
            check=False,
        )
    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        tail = "\n".join(details.splitlines()[-20:])
        raise RuntimeError(tail or f"Engine exited with code {result.returncode}")

    output_path = work_dir / "outputs" / output_name
    if not output_path.exists():
        raise RuntimeError("Generated Excel file was not found")
    return output_path


def run_job(job_id: str, job_dir: Path, env: dict[str, str], output_name: str) -> None:
    try:
        set_job(job_id, status="running", message="正在生成报表")
        output_path = run_engine(env, job_dir / "work", output_name)
        JOB_ROOT.mkdir(parents=True, exist_ok=True)
        final_path = JOB_ROOT / f"{job_id}.xlsx"
        shutil.copy2(output_path, final_path)
        set_job(job_id, status="done", message="生成完成", filename=output_name, path=str(final_path))
    except Exception as exc:
        set_job(job_id, status="failed", message="生成失败", error=str(exc))
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_ROOT / "index.html")


@app.get("/health")
def health() -> dict[str, str | int]:
    return {"status": "ok", "version": APP_VERSION[:12], "required_files": len(REQUIRED_FILES)}


@app.post("/api/upload-sessions")
def create_upload_session() -> dict[str, Any]:
    UPLOAD_SESSION_ROOT.mkdir(parents=True, exist_ok=True)
    upload_id = uuid4().hex
    session_dir = Path(tempfile.mkdtemp(prefix=upload_id + "-", dir=UPLOAD_SESSION_ROOT))
    input_dir = session_dir / "input"
    work_dir = session_dir / "work"
    (work_dir / "outputs").mkdir(parents=True, exist_ok=True)
    set_upload_session(
        upload_id,
        root_dir=str(session_dir),
        input_dir=str(input_dir),
        work_dir=str(work_dir),
        uploaded=set(),
    )
    return {"upload_id": upload_id, "required_files": len(REQUIRED_FILES)}


@app.post("/api/upload-sessions/{upload_id}/files/{key}")
async def upload_session_file(upload_id: str, key: str, request: Request) -> dict[str, Any]:
    if key not in REQUIRED_FILES:
        raise HTTPException(status_code=404, detail="unknown file key")
    session = get_upload_session(upload_id)
    form = await request.form()
    upload = form.get("file")
    if not isinstance(upload, UploadFile):
        raise HTTPException(status_code=400, detail="missing file")
    await save_upload(upload, Path(str(session["input_dir"])) / REQUIRED_FILES[key])
    with UPLOAD_SESSION_LOCK:
        current = UPLOAD_SESSIONS.get(upload_id)
        if not current:
            raise HTTPException(status_code=404, detail="upload session not found")
        uploaded = set(current.get("uploaded", set()))
        uploaded.add(key)
        current["uploaded"] = uploaded
    return {"upload_id": upload_id, "key": key, "uploaded": len(uploaded), "required_files": len(REQUIRED_FILES)}


def run_uploaded_session_job(upload_id: str, job_id: str, job_dir: Path, env: dict[str, str], output_name: str) -> None:
    try:
        run_job(job_id, job_dir, env, output_name)
    finally:
        with UPLOAD_SESSION_LOCK:
            UPLOAD_SESSIONS.pop(upload_id, None)


@app.post("/api/upload-sessions/{upload_id}/start")
async def start_upload_session_job(upload_id: str, request: Request, background_tasks: BackgroundTasks) -> dict[str, str]:
    session = get_upload_session(upload_id)
    payload = await request.json()
    current_start = require_date(payload.get("current_start"), "current_start")
    current_end = require_date(payload.get("current_end"), "current_end")
    previous_start = require_date(payload.get("previous_start"), "previous_start")
    previous_end = require_date(payload.get("previous_end"), "previous_end")

    input_dir = Path(str(session["input_dir"]))
    uploaded = set(session.get("uploaded", set()))
    missing = [key for key, name in REQUIRED_FILES.items() if key not in uploaded or not (input_dir / name).exists()]
    if missing:
        raise HTTPException(status_code=400, detail=f"缺少上传文件：{', '.join(missing)}")

    JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid4().hex
    job_dir = Path(str(session["root_dir"]))
    work_dir = Path(str(session["work_dir"]))
    env, output_name = build_env(input_dir, work_dir, current_start, current_end, previous_start, previous_end)
    set_job(job_id, status="queued", message="已上传，等待生成", filename=output_name)
    background_tasks.add_task(run_uploaded_session_job, upload_id, job_id, job_dir, env, output_name)
    return {"job_id": job_id, "status": "queued"}


@app.post("/api/jobs")
async def create_job(request: Request, background_tasks: BackgroundTasks) -> dict[str, str]:
    form = await request.form()
    current_start = require_date(form.get("current_start"), "current_start")
    current_end = require_date(form.get("current_end"), "current_end")
    previous_start = require_date(form.get("previous_start"), "previous_start")
    previous_end = require_date(form.get("previous_end"), "previous_end")

    missing = [key for key in REQUIRED_FILES if not isinstance(form.get(key), UploadFile)]
    if missing:
        raise HTTPException(status_code=400, detail=f"缺少上传文件：{', '.join(missing)}")

    JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid4().hex
    job_dir = Path(tempfile.mkdtemp(prefix=f"ora-job-{job_id}-"))
    input_dir = job_dir / "input"
    work_dir = job_dir / "work"
    (work_dir / "outputs").mkdir(parents=True, exist_ok=True)

    for key, canonical in REQUIRED_FILES.items():
        await save_upload(form[key], input_dir / canonical)

    env, output_name = build_env(input_dir, work_dir, current_start, current_end, previous_start, previous_end)
    set_job(job_id, status="queued", message="已上传，等待生成", filename=output_name)
    background_tasks.add_task(run_job, job_id, job_dir, env, output_name)
    return {"job_id": job_id, "status": "queued"}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict[str, Any]:
    job = get_job(job_id)
    return {key: value for key, value in job.items() if key != "path"}


@app.get("/api/jobs/{job_id}/download")
def job_download(job_id: str) -> FileResponse:
    job = get_job(job_id)
    if job.get("status") != "done":
        raise HTTPException(status_code=409, detail="任务尚未完成")
    path = Path(str(job.get("path", "")))
    if not path.exists():
        raise HTTPException(status_code=404, detail="生成文件不存在")
    return FileResponse(path, filename=str(job.get("filename") or path.name), media_type=EXCEL_MIME)


@app.post("/api/generate", response_model=None)
async def generate(request: Request):
    form = await request.form()
    current_start = require_date(form.get("current_start"), "current_start")
    current_end = require_date(form.get("current_end"), "current_end")
    previous_start = require_date(form.get("previous_start"), "previous_start")
    previous_end = require_date(form.get("previous_end"), "previous_end")

    missing = [key for key in REQUIRED_FILES if not isinstance(form.get(key), UploadFile)]
    if missing:
        raise HTTPException(status_code=400, detail=f"缺少上传文件：{', '.join(missing)}")

    with tempfile.TemporaryDirectory(prefix="ora-report-") as tmp:
        tmp_root = Path(tmp)
        input_dir = tmp_root / "input"
        work_dir = tmp_root / "work"
        output_dir = work_dir / "outputs"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        for key, canonical in REQUIRED_FILES.items():
            await save_upload(form[key], input_dir / canonical)

        env, output_name = build_env(input_dir, work_dir, current_start, current_end, previous_start, previous_end)
        try:
            output_path = run_engine(env, work_dir, output_name)
        except Exception as exc:
            return JSONResponse(
                status_code=500,
                content={"detail": "报表生成失败", "error": str(exc)},
            )

        final_path = Path(tempfile.gettempdir()) / f"ora-report-{uuid4().hex}.xlsx"
        shutil.copy2(output_path, final_path)

    return FileResponse(
        final_path,
        filename=output_name,
        media_type=EXCEL_MIME,
        background=BackgroundTask(lambda: final_path.unlink(missing_ok=True)),
    )
