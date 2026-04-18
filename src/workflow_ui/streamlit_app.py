from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pandas as pd
import streamlit as st
from docx import Document


st.set_page_config(page_title="YiZhi Workflow Demo", layout="wide")
st.title("YiZhi 法律工作流演示")


def get_json(base_url: str, path: str) -> dict[str, Any]:
    url = f"{base_url.rstrip('/')}{path}"
    response = httpx.get(url, timeout=15.0)
    if response.is_error:
        raise RuntimeError(f"GET {path} failed: {response.status_code} {response.text}")
    return response.json()


def post_json(base_url: str, path: str, payload: dict[str, Any]) -> dict[str, Any]:
    url = f"{base_url.rstrip('/')}{path}"
    response = httpx.post(url, json=payload, timeout=300.0)
    if response.is_error:
        raise RuntimeError(f"POST {path} failed: {response.status_code} {response.text}")
    return response.json()


def read_excel_preview(file_path: Path, max_rows: int = 30) -> pd.DataFrame:
    return pd.read_excel(file_path).head(max_rows)


def read_docx_preview(file_path: Path, max_paragraphs: int = 20) -> list[str]:
    doc = Document(file_path)
    paragraphs: list[str] = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        paragraphs.append(text)
        if len(paragraphs) >= max_paragraphs:
            break
    return paragraphs


with st.sidebar:
    st.subheader("后端配置")
    backend_base_url = st.text_input("API Base URL", value="http://127.0.0.1:8000")
    if st.button("健康检查", use_container_width=True):
        try:
            health = get_json(backend_base_url, "/healthz")
            st.success(f"后端可用: {health}")
        except Exception as exc:  # noqa: BLE001
            st.error(f"健康检查失败: {exc}")

tabs = st.tabs(["Excel 工作流", "Word 工作流"])

with tabs[0]:
    st.subheader("生成 Excel")
    excel_query = st.text_area("检索问题", height=100, value="上下班途中交通事故工伤案例")
    excel_target_case_count = st.number_input(
        "目标案例数量",
        min_value=1,
        max_value=200,
        value=50,
        step=1,
    )

    if st.button("运行 Excel 工作流", type="primary", use_container_width=True):
        payload = {
            "query": excel_query.strip(),
            "target_case_count": int(excel_target_case_count),
        }
        try:
            result = post_json(backend_base_url, "/api/workflows/excel/run", payload)
            st.session_state["excel_result"] = result
            st.success("Excel 工作流执行完成")
        except Exception as exc:  # noqa: BLE001
            st.error(f"执行失败: {exc}")

    excel_result = st.session_state.get("excel_result")
    if excel_result:
        st.json(excel_result)
        excel_path = Path(excel_result["output_path"])
        if excel_path.exists():
            preview_df = read_excel_preview(excel_path)
            st.dataframe(preview_df, use_container_width=True)
            st.download_button(
                label="下载 Excel",
                data=excel_path.read_bytes(),
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning(f"文件不存在: {excel_path}")

with tabs[1]:
    st.subheader("生成 Word 报告")
    word_query = st.text_area("检索问题", height=100, value="上下班途中交通事故工伤案例", key="word_query")
    word_submitter = st.text_input("提交人", value="xxxxxx")
    word_report_title = st.text_input("报告标题", value="案涉争议问题检索报告")
    word_report_date = st.text_input("报告日期（可选）", value="")

    if st.button("运行 Word 工作流", type="primary", use_container_width=True):
        payload: dict[str, Any] = {
            "query": word_query.strip(),
            "submitter": word_submitter.strip(),
            "report_title": word_report_title.strip(),
        }
        if word_report_date.strip():
            payload["report_date"] = word_report_date.strip()

        try:
            result = post_json(backend_base_url, "/api/workflows/word/run", payload)
            st.session_state["word_result"] = result
            st.success("Word 工作流执行完成")
        except Exception as exc:  # noqa: BLE001
            st.error(f"执行失败: {exc}")

    word_result = st.session_state.get("word_result")
    if word_result:
        st.json(word_result)
        word_path = Path(word_result["output_path"])
        if word_path.exists():
            preview_paragraphs = read_docx_preview(word_path)
            st.markdown("**文档预览（前20段）**")
            st.text("\n\n".join(preview_paragraphs))
            st.download_button(
                label="下载 Word",
                data=word_path.read_bytes(),
                file_name=word_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.warning(f"文件不存在: {word_path}")
