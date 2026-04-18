# Streamlit 演示页

## 1. 启动后端 API

```powershell
conda run -n YiZhi python -m uvicorn workflow_api.app:app --host 127.0.0.1 --port 8000
```

## 2. 启动 Streamlit

```powershell
conda run -n YiZhi python -m streamlit run src/workflow_ui/streamlit_app.py
```

访问 `http://127.0.0.1:8501`

## 3. 使用说明

- 侧边栏填后端地址（默认 `http://127.0.0.1:8000`），先点“健康检查”。
- `Excel 工作流` 页可提交 `query + target_case_count`，支持结果预览和 `.xlsx` 下载。
- `Word 工作流` 页可提交 `query + submitter + report_title + report_date`，支持正文预览和 `.docx` 下载。
