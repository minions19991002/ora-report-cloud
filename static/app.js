const REQUIRED_FILES = [
  { key: "previous", title: "上一期报表", canonical: "上一期报表.xlsx", match: [/外送(?:周报|月报)|周报|月报/i] },
  { key: "template", title: "报表模板", canonical: "ora外送报表模板.xlsx", match: [/模板|模版/i, /外送(?:周报|月报)|周报|月报/i] },
  { key: "storeInfo", title: "ORA门店信息表", canonical: "ORA门店信息表.xlsx", match: [/门店信息/i] },
  { key: "mtStore", title: "美团门店数据", canonical: "美团门店数据.xlsx", match: [/美团门店数据/i] },
  { key: "eleStore", title: "饿了么门店数据", canonical: "饿了么门店数据.xlsx", match: [/饿了么门店数据/i] },
  { key: "mtPromo", title: "美团推广", canonical: "美团推广.xlsx", match: [/美团推广/i] },
  { key: "elePromo", title: "饿了么推广", canonical: "饿了么推广.xlsx", match: [/饿了么推广/i] },
  { key: "mtOrder", title: "美团订单数据", canonical: "美团订单数据.xlsx", match: [/美团订单/i] },
  { key: "eleOrder", title: "饿了么订单数据", canonical: "饿了么订单数据.xlsx", match: [/饿了么订单/i] },
  { key: "distance", title: "订单距离分布", canonical: "订单距离分布.xlsx", match: [/订单距离|距离分布/i] },
  { key: "mtProduct", title: "美团商品数据", canonical: "美团商品数据.xlsx", match: [/美团商品/i] },
  { key: "eleProduct", title: "饿了么商品数据", canonical: "饿了么商品数据.xlsx", match: [/饿了么商品/i] },
  { key: "reviewSummary", title: "评价汇总", canonical: "评价汇总.xlsx", match: [/评价汇总/i] },
  { key: "reviewCounts", title: "好评数中差评", canonical: "好评数中差评数据.xlsx", match: [/好评数|中差评数据/i] },
  { key: "delivery", title: "美团平均配送时长", canonical: "美团平均配送时长.xlsx", match: [/平均配送时长|配送时长/i] },
  { key: "oraDaily", title: "ORA外送日报", canonical: "Ora外送日报.xlsx", match: [/ora\s*外送(?:日报|周报)|外送日报/i] },
  { key: "oraProduct", title: "ORA外送商品数据", canonical: "Ora外送商品数据.xlsx", match: [/ora外送商品数据|外送商品数据/i] },
];

const HEADER_SIGNATURES = {
  previous: {
    all: ["店号", "店名", "Sales_日均", "总sales", "订单距离分布"],
    sheetAll: ["V2", "上期", "订单距离及实付区间", "商品销售排行-单品", "商品销售排行-套餐", "用户体验-客诉", "用户体验-配送"],
    sheetAny: ["源数据", "透视", "商品销售数据源", "商品销售透视", "本周业绩预测"],
    sheetAnyRequired: true,
  },
  template: {
    all: ["店号", "店名", "Sales_日均", "总sales", "订单距离分布"],
    sheetAll: ["V2", "上期", "订单距离及实付区间", "商品销售排行-单品", "商品销售排行-套餐", "用户体验-客诉", "用户体验-配送"],
    sheetNot: ["源数据", "透视", "商品销售数据源", "商品销售透视", "本周业绩预测", "草稿"],
  },
  storeInfo: { all: ["店号", "店名"], any: ["美团门店ID", "饿了么门店ID"] },
  mtStore: { all: ["营业收入", "平台服务费", "商家活动支出", "有效订单"], any: ["曝光提升数(次)", "综合体验分"] },
  eleStore: { all: ["收入", "平台技术服务费", "履约技术服务费", "有效订单"], any: ["商家活动成本（含满减活动）", "店铺评分"] },
  mtPromo: { all: ["推广消费实付(元)", "访问提升数(次)"], anyAll: [["订单交易额", "订单原价交易额", "推广营业额"]], sheetAny: ["效果数据"] },
  elePromo: { all: ["推广现金消费(元)", "进店提升数"], anyAll: [["订单交易额", "订单原价交易额", "推广营业额"]] },
  mtOrder: { all: ["订单实付"], any: ["订单状态", "已完成", "已接单"] },
  eleOrder: { all: ["顾客实付"], any: ["订单完结", "是否预订单", "接单时间", "完成时间"] },
  distance: { all: ["[0,0.5Km)", "[0.5,0.8Km)", "[0.8,1.0Km)"], sheetAny: ["订单分布"] },
  mtProduct: { all: ["销量", "销售额"], any: ["商品名", "商品名称", "菜品名称", "规格"], not: ["是否套餐"] },
  eleProduct: { all: ["销量", "销售额"], any: ["商品名称", "菜品名称", "是否套餐"], sheetAny: ["data"] },
  reviewSummary: { all: ["平台", "一级分类", "二级分类", "统计"], sheetAny: ["双平台中差评汇总", "数据源"] },
  reviewCounts: { all: ["好评数", "中差评数"] },
  delivery: { all: ["平均配送时长"], any: ["美团门店ID", "门店ID", "门店名称"] },
  oraDaily: {
    all: ["store_id", "sales_channel"],
    anyAll: [["gross_amount", "销售额", "总sales"], ["order_count", "订单数", "有效订单", "ADT"], ["discount_amount", "优惠金额", "商户折扣金额"]],
    sheetAny: ["Ora外送日报", "Ora 外送日报", "Ora外送周报", "Ora 外送周报"],
  },
  oraProduct: { all: ["date_id", "sku_name", "store_id", "sales_channel", "quantity", "gross_amount"], sheetAny: ["Ora外送商品数据"] },
};

const XLSX_SCRIPT_URL = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
const UPLOAD_CONCURRENCY = 2;

const filesByKey = new Map();
const grid = document.querySelector("#fileGrid");
const logEl = document.querySelector("#log");
const bulkInput = document.querySelector("#bulkInput");
const dropzone = document.querySelector("#dropzone");
const currentStartDate = document.querySelector("#currentStartDate");
const currentEndDate = document.querySelector("#currentEndDate");
const previousStartDate = document.querySelector("#previousStartDate");
const previousEndDate = document.querySelector("#previousEndDate");
const periodText = document.querySelector("#periodText");
const readyCount = document.querySelector("#readyCount");
const statusText = document.querySelector("#statusText");
const generateBtn = document.querySelector("#generateBtn");
const clearBtn = document.querySelector("#clearBtn");
const progress = document.querySelector("#progress");

let xlsxPromise = null;
let xlsxWarningShown = false;

function log(message) {
  const stamp = new Date().toLocaleTimeString("zh-CN", { hour12: false });
  logEl.textContent += `[${stamp}] ${message}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function loadExternalScript(url, globalName) {
  if (globalThis[globalName]) return Promise.resolve(globalThis[globalName]);
  return new Promise((resolve, reject) => {
    const existing = [...document.scripts].find((script) => script.src === url);
    if (existing) {
      existing.addEventListener("load", () => resolve(globalThis[globalName]));
      existing.addEventListener("error", () => reject(new Error(`${url} 加载失败`)));
      return;
    }
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.onload = () => resolve(globalThis[globalName]);
    script.onerror = () => reject(new Error(`${url} 加载失败`));
    document.head.appendChild(script);
  });
}

async function ensureXlsxRuntime() {
  if (!xlsxPromise) xlsxPromise = loadExternalScript(XLSX_SCRIPT_URL, "XLSX");
  try {
    await xlsxPromise;
    return Boolean(globalThis.XLSX);
  } catch {
    if (!xlsxWarningShown) {
      log("表头识别库加载失败，已改用文件名匹配。");
      xlsxWarningShown = true;
    }
    return false;
  }
}

function classifyByName(file) {
  const scored = REQUIRED_FILES.map((item) => ({
    item,
    score: item.match.reduce((sum, pattern) => sum + (pattern.test(file.name) ? 1 : 0), 0),
    maxScore: item.match.length,
  })).filter((item) => item.score > 0);
  scored.sort((a, b) => b.score - a.score);
  return scored[0]?.item ? { item: scored[0].item, method: "文件名", score: scored[0].score, maxScore: scored[0].maxScore } : null;
}

function normalizeHint(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）]/g, "");
}

function hasHint(hints, term) {
  const needle = normalizeHint(term);
  return hints.text.includes(needle) || hints.sheets.some((sheet) => sheet.includes(needle));
}

function hasSheetHint(hints, term) {
  const needle = normalizeHint(term);
  return hints.sheets.some((sheet) => sheet.includes(needle));
}

async function readWorkbookHints(file) {
  if (!globalThis.XLSX && !(await ensureXlsxRuntime())) return null;
  const workbook = globalThis.XLSX.read(await file.arrayBuffer(), { type: "array", sheetRows: 24, cellDates: false });
  const sheets = workbook.SheetNames.map(normalizeHint);
  const cells = [];
  for (const sheetName of workbook.SheetNames.slice(0, 12)) {
    const rows = globalThis.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });
    for (const row of rows.slice(0, 24)) {
      for (const cell of row.slice(0, 80)) {
        const text = normalizeHint(cell);
        if (text) cells.push(text);
      }
    }
  }
  return { sheets, text: cells.join("|") };
}

function signatureScore(hints, signature) {
  if (!signature) return 0;
  let score = 0;

  for (const term of signature.all ?? []) {
    if (!hasHint(hints, term)) return 0;
    score += 6;
  }
  for (const term of signature.sheetAll ?? []) {
    if (!hasSheetHint(hints, term)) return 0;
    score += 5;
  }

  const anyTerms = signature.any ?? [];
  const anyHits = anyTerms.filter((term) => hasHint(hints, term)).length;
  if (anyTerms.length > 0 && anyHits === 0) return 0;
  score += anyHits * 3;

  for (const group of signature.anyAll ?? []) {
    const groupHits = group.filter((term) => hasHint(hints, term)).length;
    if (groupHits === 0) return 0;
    score += groupHits * 4;
  }

  const sheetTerms = signature.sheetAny ?? [];
  const sheetHits = sheetTerms.filter((term) => hasSheetHint(hints, term)).length;
  if (sheetTerms.length > 0 && sheetHits === 0 && ((signature.all?.length ?? 0) === 0 || signature.sheetAnyRequired)) return 0;
  score += sheetHits * 4;

  for (const term of signature.not ?? []) {
    if (hasHint(hints, term)) score -= 8;
  }
  for (const term of signature.sheetNot ?? []) {
    if (hasSheetHint(hints, term)) return 0;
  }

  return Math.max(score, 0);
}

async function classifyByHeader(file) {
  try {
    const hints = await readWorkbookHints(file);
    if (!hints) return null;
    const scored = REQUIRED_FILES.map((item) => ({
      item,
      score: signatureScore(hints, HEADER_SIGNATURES[item.key]),
    })).filter((item) => item.score > 0);
    scored.sort((a, b) => b.score - a.score);
    return scored[0]?.score >= 6 ? { item: scored[0].item, method: "表头内容" } : null;
  } catch {
    log(`表头识别失败，改用文件名匹配：${file.name}`);
    return null;
  }
}

async function classify(file) {
  const nameResult = classifyByName(file);
  const ambiguousReportName = ["previous", "template"].includes(nameResult?.item?.key);
  if (nameResult?.score >= nameResult?.maxScore && !ambiguousReportName) {
    return nameResult;
  }
  const headerResult = await classifyByHeader(file);
  const nameKey = nameResult?.item?.key;
  const headerKey = headerResult?.item?.key;
  if (
    nameKey &&
    headerKey &&
    nameKey !== headerKey &&
    ["oraDaily", "oraProduct"].includes(headerKey)
  ) {
    return headerResult;
  }
  if (
    nameKey &&
    headerKey &&
    nameKey !== headerKey &&
    ["previous", "template"].includes(nameKey) &&
    ["previous", "template"].includes(headerKey)
  ) {
    return nameResult;
  }
  return headerResult ?? nameResult;
}

function isIgnoredFile(name) {
  return /最新评分/i.test(name);
}

function parseDatesFromName(name) {
  const fullDates = [...name.matchAll(/20\d{2}[-._]\d{1,2}[-._]\d{1,2}/g)].map((match) => match[0].replace(/[._]/g, "-"));
  const compactRanges = [...name.matchAll(/(?<!\d)(\d{1,2})\.(\d{1,2})[-~](\d{1,2})\.(\d{1,2})(?!\d)/g)].flatMap((match) => {
    const year = new Date().getFullYear();
    return [`${year}-${match[1]}-${match[2]}`, `${year}-${match[3]}-${match[4]}`];
  });
  const monthRanges = [...name.matchAll(/(?:(20\d{2})[年._-]?)?(\d{1,2})月/g)].flatMap((match) => {
    const year = Number(match[1] || new Date().getFullYear());
    const month = Number(match[2]);
    if (month < 1 || month > 12) return [];
    const lastDay = new Date(year, month, 0).getDate();
    return [`${year}-${month}-1`, `${year}-${month}-${lastDay}`];
  });
  return [...new Set([...fullDates, ...compactRanges, ...monthRanges])].map((date) => {
    const [year, month, day] = date.split("-").map(Number);
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  });
}

function labelFromDates(start, end) {
  const startDate = new Date(`${start}T00:00:00`);
  const endDate = new Date(`${end}T00:00:00`);
  const lastDay = new Date(endDate.getFullYear(), endDate.getMonth() + 1, 0).getDate();
  const isFullMonthRange = startDate.getDate() === 1 && endDate.getDate() === lastDay && endDate >= startDate;
  if (isFullMonthRange) {
    const startMonth = startDate.getMonth() + 1;
    const endMonth = endDate.getMonth() + 1;
    if (startDate.getFullYear() === endDate.getFullYear() && startMonth === endMonth) return `${startMonth}月`;
    if (startDate.getFullYear() === endDate.getFullYear()) return `${startMonth}-${endMonth}月`;
    return `${startDate.getFullYear()}.${startMonth}-${endDate.getFullYear()}.${endMonth}月`;
  }
  return `${startDate.getMonth() + 1}.${startDate.getDate()}-${endDate.getMonth() + 1}.${endDate.getDate()}`;
}

function exportLabelFromDates(start, end) {
  const startDate = new Date(`${start}T00:00:00`);
  const endDate = new Date(`${end}T00:00:00`);
  const lastDay = new Date(endDate.getFullYear(), endDate.getMonth() + 1, 0).getDate();
  const isFullMonthRange = startDate.getDate() === 1 && endDate.getDate() === lastDay && endDate >= startDate;
  if (isFullMonthRange) {
    const startMonth = startDate.getMonth() + 1;
    const endMonth = endDate.getMonth() + 1;
    if (startDate.getFullYear() === endDate.getFullYear() && startMonth === endMonth) return `${startMonth}月`;
    if (startDate.getFullYear() === endDate.getFullYear()) return `${startMonth}-${endMonth}月`;
    return `${startDate.getFullYear()}.${startMonth}-${endDate.getFullYear()}.${endMonth}月`;
  }
  return labelFromDates(start, end);
}

function inferPreviousPeriod() {
  if (!currentStartDate.value || !currentEndDate.value || previousStartDate.value || previousEndDate.value) return;
  const currentStart = new Date(`${currentStartDate.value}T00:00:00`);
  const currentEnd = new Date(`${currentEndDate.value}T00:00:00`);
  if (Number.isNaN(currentStart.getTime()) || Number.isNaN(currentEnd.getTime())) return;
  const periodDays = Math.round((currentEnd - currentStart) / 86400000) + 1;
  const prevEnd = new Date(currentStart);
  prevEnd.setDate(currentStart.getDate() - 1);
  const prevStart = new Date(prevEnd);
  prevStart.setDate(prevEnd.getDate() - periodDays + 1);
  previousStartDate.value = prevStart.toISOString().slice(0, 10);
  previousEndDate.value = prevEnd.toISOString().slice(0, 10);
}

function updatePeriod() {
  if (currentStartDate.value && currentEndDate.value && previousStartDate.value && previousEndDate.value) {
    periodText.textContent = `本期 ${labelFromDates(currentStartDate.value, currentEndDate.value)} ｜ 上期 ${labelFromDates(previousStartDate.value, previousEndDate.value)}`;
  } else if (currentStartDate.value && currentEndDate.value) {
    periodText.textContent = `本期 ${labelFromDates(currentStartDate.value, currentEndDate.value)} ｜ 请补充上期周期`;
  } else {
    periodText.textContent = "待识别周期";
  }
}

function autoFillDates(files) {
  const candidates = [...new Set(files.flatMap((file) => parseDatesFromName(file.name)))].sort();
  if (candidates.length >= 4) {
    previousStartDate.value ||= candidates[0];
    previousEndDate.value ||= candidates[1];
    currentStartDate.value ||= candidates[candidates.length - 2];
    currentEndDate.value ||= candidates[candidates.length - 1];
  } else if (candidates.length >= 2) {
    currentStartDate.value ||= candidates[0];
    currentEndDate.value ||= candidates[candidates.length - 1];
    inferPreviousPeriod();
  } else if (candidates.length === 1 && !currentEndDate.value) {
    const end = new Date(`${candidates[0]}T00:00:00`);
    const start = new Date(end);
    start.setDate(end.getDate() - 6);
    currentStartDate.value = start.toISOString().slice(0, 10);
    currentEndDate.value = candidates[0];
    inferPreviousPeriod();
  }
  updatePeriod();
}

async function handleFiles(fileList) {
  const files = [...fileList];
  autoFillDates(files);
  let matched = 0;
  for (const file of files) {
    if (isIgnoredFile(file.name)) {
      log(`已忽略：${file.name}（最新评分表不需要上传）`);
      continue;
    }
    const result = await classify(file);
    if (!result) {
      log(`未识别：${file.name}`);
      continue;
    }
    filesByKey.set(result.item.key, file);
    matched += 1;
    log(`已按${result.method}匹配：${file.name} → ${result.item.title}`);
  }
  if (matched === 0 && files.length) log("没有匹配到标准源表。");
  renderSlots();
}

function renderSlots() {
  grid.innerHTML = "";
  let ready = 0;
  for (const item of REQUIRED_FILES) {
    const file = filesByKey.get(item.key);
    if (file) ready += 1;
    const slot = document.createElement("article");
    slot.className = `slot ${file ? "ready" : "missing"}`;
    slot.innerHTML = `
      <div class="slotHeader">
        <div class="slotTitle">${item.title}</div>
        <span class="badge ${file ? "ready" : "missing"}">${file ? "已上传" : "缺失"}</span>
      </div>
      <div class="filename">${file ? file.name : item.canonical}</div>
      <div class="slotActions">
        <button type="button">选择文件</button>
        <input type="file" accept=".xlsx,.xls" />
      </div>
    `;
    const input = slot.querySelector("input");
    slot.querySelector("button").addEventListener("click", () => input.click());
    input.addEventListener("change", (event) => {
      const selected = event.target.files?.[0];
      if (!selected) return;
      filesByKey.set(item.key, selected);
      autoFillDates([selected]);
      log(`已上传：${item.title} ← ${selected.name}`);
      renderSlots();
    });
    grid.appendChild(slot);
  }
  readyCount.textContent = `${ready}/${REQUIRED_FILES.length}`;
  statusText.textContent = ready === REQUIRED_FILES.length ? "源表齐全" : "源表缺失";
  generateBtn.disabled = ready !== REQUIRED_FILES.length;
}

function filenameFromDisposition(disposition, fallback) {
  const utf8 = disposition?.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) return decodeURIComponent(utf8[1]);
  const ascii = disposition?.match(/filename="?([^"]+)"?/i);
  return ascii?.[1] || fallback;
}

async function pollJob(jobId) {
  const startedAt = Date.now();
  let lastRunningLogAt = 0;
  for (;;) {
    const response = await fetch(`/api/jobs/${jobId}`);
    if (!response.ok) throw new Error(await response.text());
    const job = await response.json();
    const elapsedSeconds = Math.round((Date.now() - startedAt) / 1000);
    if (job.status === "queued") {
      progress.value = Math.max(progress.value, 42);
      statusText.textContent = "排队中";
    } else if (job.status === "running") {
      progress.value = Math.max(progress.value, Math.min(88, 65 + Math.floor(elapsedSeconds / 10)));
      statusText.textContent = `服务器生成中 ${Math.floor(elapsedSeconds / 60)}:${String(elapsedSeconds % 60).padStart(2, "0")}`;
      if (elapsedSeconds - lastRunningLogAt >= 30) {
        log(`服务器仍在生成，已用 ${elapsedSeconds} 秒；一般需要 2-4 分钟。`);
        lastRunningLogAt = elapsedSeconds;
      }
    } else if (job.status === "done") {
      progress.value = 90;
      statusText.textContent = "下载中";
      return job;
    } else if (job.status === "failed") {
      throw new Error(job.error || job.message || "生成失败");
    }
    await new Promise((resolve) => window.setTimeout(resolve, 2000));
  }
}

async function uploadFileToSession(uploadId, item, completed, total) {
  const file = filesByKey.get(item.key);
  const form = new FormData();
  form.append("file", file, file.name);
  const response = await fetch(`/api/upload-sessions/${uploadId}/files/${item.key}`, {
    method: "POST",
    body: form,
  });
  if (!response.ok) throw new Error(await response.text());
  const uploaded = completed();
  progress.value = Math.max(progress.value, 12 + Math.floor((uploaded / total) * 43));
  statusText.textContent = `上传中 ${uploaded}/${total}`;
  log(`已上传 ${uploaded}/${total}：${item.title}`);
}

async function uploadFileToSessionWithRetry(uploadId, item, completed, total) {
  const maxAttempts = 5;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      await uploadFileToSession(uploadId, item, completed, total);
      return;
    } catch (error) {
      if (attempt === maxAttempts) {
        throw new Error(`${item.title} 上传失败：${error?.message || String(error)}`);
      }
      const delaySeconds = attempt * 5;
      log(`${item.title} 上传中断，${delaySeconds} 秒后重试（${attempt + 1}/${maxAttempts}）`);
      await new Promise((resolve) => window.setTimeout(resolve, delaySeconds * 1000));
    }
  }
}

async function uploadFilesToSession() {
  const sessionResponse = await fetch("/api/upload-sessions", { method: "POST" });
  if (!sessionResponse.ok) throw new Error(await sessionResponse.text());
  const session = await sessionResponse.json();
  const uploadId = session.upload_id;
  let uploaded = 0;
  const nextCompleted = () => {
    uploaded += 1;
    return uploaded;
  };
  const queue = [...REQUIRED_FILES].sort((a, b) => (filesByKey.get(b.key)?.size || 0) - (filesByKey.get(a.key)?.size || 0));
  const totalUploads = REQUIRED_FILES.length;
  const workerCount = Math.min(UPLOAD_CONCURRENCY, queue.length);
  await Promise.all(
    Array.from({ length: workerCount }, async () => {
      while (queue.length) {
        const item = queue.shift();
        await uploadFileToSessionWithRetry(uploadId, item, nextCompleted, totalUploads);
      }
    }),
  );
  return uploadId;
}

async function generateReport() {
  if (!currentStartDate.value || !currentEndDate.value || !previousStartDate.value || !previousEndDate.value) {
    statusText.textContent = "周期缺失";
    return;
  }
  const readyRequired = REQUIRED_FILES.filter((item) => filesByKey.has(item.key)).length;
  if (readyRequired !== REQUIRED_FILES.length) {
    statusText.textContent = "源表缺失";
    return;
  }

  generateBtn.disabled = true;
  clearBtn.disabled = true;
  progress.value = 12;
  statusText.textContent = "上传生成中";
  log("开始分文件上传到服务器...");

  try {
    const uploadId = await uploadFilesToSession();
    progress.value = 56;
    statusText.textContent = "创建生成任务";
    log("文件上传完成，开始创建服务器生成任务...");
    const response = await fetch(`/api/upload-sessions/${uploadId}/start`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        current_start: currentStartDate.value,
        current_end: currentEndDate.value,
        previous_start: previousStartDate.value,
        previous_end: previousEndDate.value,
      }),
    });
    if (!response.ok) throw new Error(await response.text());
    const created = await response.json();
    log(`服务器任务已创建：${created.job_id}`);
    const job = await pollJob(created.job_id);
    const downloadResponse = await fetch(`/api/jobs/${created.job_id}/download`);
    if (!downloadResponse.ok) throw new Error(await downloadResponse.text());
    const blob = await downloadResponse.blob();
    const fallbackName = `ORA外送报表_${exportLabelFromDates(currentStartDate.value, currentEndDate.value)}.xlsx`;
    const filename = filenameFromDisposition(downloadResponse.headers.get("Content-Disposition"), job.filename || fallbackName);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    URL.revokeObjectURL(url);

    progress.value = 100;
    statusText.textContent = "已生成";
    log(`已下载：${filename}`);
  } catch (error) {
    statusText.textContent = "生成失败";
    const message = error?.message || String(error);
    if (/Failed to fetch/i.test(message)) {
      log("无法连接服务器：云端服务当前可用，但本次上传连接被浏览器或网络中断。请刷新页面后重试；如果在公司网络/代理环境下仍失败，换一个网络或让文件上传者直接在本机操作。");
    } else {
      log(message);
    }
  } finally {
    clearBtn.disabled = false;
    generateBtn.disabled = REQUIRED_FILES.filter((item) => filesByKey.has(item.key)).length !== REQUIRED_FILES.length;
  }
}

bulkInput.addEventListener("change", (event) => {
  void handleFiles(event.target.files);
});
currentStartDate.addEventListener("change", () => {
  inferPreviousPeriod();
  updatePeriod();
});
currentEndDate.addEventListener("change", () => {
  inferPreviousPeriod();
  updatePeriod();
});
previousStartDate.addEventListener("change", updatePeriod);
previousEndDate.addEventListener("change", updatePeriod);
generateBtn.addEventListener("click", generateReport);
clearBtn.addEventListener("click", () => {
  filesByKey.clear();
  logEl.textContent = "";
  progress.value = 0;
  currentStartDate.value = "";
  currentEndDate.value = "";
  previousStartDate.value = "";
  previousEndDate.value = "";
  updatePeriod();
  renderSlots();
});

["dragenter", "dragover"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.add("active");
  });
});

["dragleave", "drop"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.remove("active");
  });
});

dropzone.addEventListener("drop", (event) => {
  void handleFiles(event.dataTransfer.files);
});

renderSlots();
