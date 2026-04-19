# YiZhi

异智（YiZhi）是一个面向法律场景的类案检索与报告生成项目。  
项目提供两类核心能力：

- 生成类案检索 Excel（结构化案件字段）
- 生成类案检索 Word 报告（基于检索结果与提示词）

---

## 1. 环境要求

- Python `>=3.11`
- 可访问：DeliLegal API + 你的 LLM API

---

## 2. 安装

```bash
git clone https://github.com/chuoer47/YiZhi.git
cd YiZhi
pip install -e .
```

可选（conda）：

```bash
conda create -n YiZhi python=3.11 -y
conda activate YiZhi
pip install -e .
```

---

## 3. 配置 `.env`

先复制模板：

```bash
cp .env.example .env
```

Windows PowerShell 可用：

```powershell
Copy-Item .env.example .env
```

然后填写以下变量：

- `DELILEGAL_APP_ID`
- `DELILEGAL_SECRET`
- `LLM_API_KEY`
- `LLM_BASE_URL`
- `LLM_MODEL`

可选：

- `DELILEGAL_BASE_URL`（默认 `https://openapi.delilegal.com`）
- `DELILEGAL_TIMEOUT`（默认 `30`）

---

## 4. 启动后端 API

```bash
python -m uvicorn workflow_api.app:app --host 127.0.0.1 --port 8000
```

启动后可访问：

- Swagger: `http://127.0.0.1:8000/docs`
- 健康检查: `http://127.0.0.1:8000/healthz`

---

## 5. 启动前端（Streamlit）

```bash
python -m streamlit run src/workflow_ui/streamlit_app.py
```

访问：`http://127.0.0.1:8501`

使用流程：

1. 在页面中确认后端地址（默认 `http://127.0.0.1:8000`）并做健康检查  
2. 输入检索问题  
3. 生成 Excel / Word（或一键生成）  
4. 在页面下载生成文件

---

## 6. API 快速测试（可选）

### Excel

`POST /api/workflows/excel/run`

```json
{
  "query": "上下班途中交通事故工伤案例",
  "target_case_count": 5
}
```

### Word

`POST /api/workflows/word/run`

```json
{
  "query": "上下班途中交通事故工伤案例",
  "submitter": "xxxxxx",
  "report_title": "案涉争议问题检索报告"
}
```

---

## 7. 常见问题

- 健康检查正常但生成失败：通常是 LLM 侧问题（余额不足、模型不支持结构化输出、API 权限不足）  
- 生成成功但看不到文件：输出默认在项目根目录 `outputs/`
