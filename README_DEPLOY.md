# ORA 外送周报云端生成版

这是服务器端生成版。浏览器只负责上传 15 个 Excel 和下载结果，Excel 处理在服务器 Python 环境中运行。

生成采用后台任务模式：上传完成后立即创建任务，页面轮询任务状态，完成后下载 Excel，避免长请求被云平台超时中断。

## 本地运行

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn server.app:app --host 0.0.0.0 --port 8000
```

打开：

```text
http://localhost:8000
```

## Docker 运行

```bash
docker build -t ora-report-cloud .
docker run --rm -p 8000:8000 ora-report-cloud
```

打开：

```text
http://localhost:8000
```

## 云端部署建议

推荐部署到支持 Docker / Python Web 服务的平台，例如 Render、Railway、Fly.io、阿里云、腾讯云或公司内网服务器。

部署参数：

- Build command: 使用 Dockerfile 自动构建，或 `pip install -r requirements.txt`
- Start command: `uvicorn server.app:app --host 0.0.0.0 --port $PORT`
- Python version: 3.12

## 使用说明

需要上传 15 个表格：

1. 上一期周报
2. 周报模板
3. ORA门店信息表
4. 美团门店数据
5. 饿了么门店数据
6. 美团推广
7. 饿了么推广
8. 美团订单数据
9. 饿了么订单数据
10. 订单距离分布
11. 美团商品数据
12. 饿了么商品数据
13. 评价汇总
14. 好评数中差评
15. 美团平均配送时长

最新评分表不需要上传。

## 数据安全

上传的 Excel 会在服务器临时目录中处理，生成完成后临时目录会释放。若表格包含敏感经营数据，建议部署在公司自有服务器、公司云账号或内网环境中。
