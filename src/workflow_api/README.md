# workflow_api 后端启动说明

## 1. 功能

`workflow_api` 是当前项目的 HTTP 后端，基于 FastAPI，对外暴露：

- `GET /healthz`：健康检查
- `POST /api/workflows/excel/run`：生成类案检索 Excel
- `POST /api/workflows/word/run`：生成类案检索 Word 报告

默认启动地址：`http://127.0.0.1:8000`

---

## 2. 启动方式

在项目根目录（`F:\code_lib\YiZhi`）执行：

### 方式 A（推荐）

```powershell
conda activate YiZhi
python -m uvicorn workflow_api.app:app --host 127.0.0.1 --port 8000
```

### 方式 B（单命令）

```powershell
conda run -n YiZhi python -m uvicorn workflow_api.app:app --host 127.0.0.1 --port 8000
```

启动后可访问：

- 文档页：`http://127.0.0.1:8000/docs`
- 健康检查：`http://127.0.0.1:8000/healthz`

---

## 3. 快速验证

### 健康检查

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:8000/healthz" -Method Get
```

### 触发 Excel 工作流

```powershell
$payload = @{
  query = "上下班途中交通事故工伤案例"
  target_case_count = 5
} | ConvertTo-Json

Invoke-RestMethod `
  -Uri "http://127.0.0.1:8000/api/workflows/excel/run" `
  -Method Post `
  -ContentType "application/json" `
  -Body $payload
```

### 触发 Word 工作流

```powershell
$payload = @{
  query = "上下班途中交通事故工伤案例"
  submitter = "xxxxxx"
  report_title = "案涉争议问题检索报告"
} | ConvertTo-Json

Invoke-RestMethod `
  -Uri "http://127.0.0.1:8000/api/workflows/word/run" `
  -Method Post `
  -ContentType "application/json" `
  -Body $payload
```

---

## 4. 目录说明

- [app.py](./app.py)：FastAPI 入口与路由定义
- [__init__.py](./__init__.py)：包标记文件
