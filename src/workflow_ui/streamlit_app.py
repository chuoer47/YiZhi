from __future__ import annotations

from collections import Counter
from io import BytesIO
from pathlib import Path
from typing import Any

import httpx
import pandas as pd
import streamlit as st
from docx import Document


st.set_page_config(
    page_title="类案智裁Agent",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_styles() -> None:
    st.markdown(
        """
<style>
:root {
  --primary: #586b8b;
  --text: #3c4252;
  --text-light: #7e8594;
  --bg: #f7f8fa;
  --card: #ffffff;
  --border: #e8ebf0;
  --hover: #f1f3f6;
}
.stApp {
  background: var(--bg);
}
.header-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 18px 20px;
  margin-bottom: 14px;
}
.app-title {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 24px;
  font-weight: 700;
  color: var(--text);
}
.search-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 18px;
  margin-bottom: 14px;
}
.panel-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 16px;
  margin-bottom: 14px;
}
.panel-title {
  font-size: 16px;
  font-weight: 700;
  color: var(--text);
  margin-bottom: 8px;
}
.muted {
  color: var(--text-light);
  font-size: 13px;
}
.law-item {
  padding: 10px 0;
  border-bottom: 1px solid var(--border);
}
.law-item:last-child {
  border-bottom: none;
}
</style>
""",
        unsafe_allow_html=True,
    )


def init_state() -> None:
    defaults: dict[str, Any] = {
        "excel_result": None,
        "word_result": None,
        "excel_df": None,
        "word_preview": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def get_json(base_url: str, path: str) -> dict[str, Any]:
    url = f"{base_url.rstrip('/')}{path}"
    response = httpx.get(url, timeout=20.0)
    if response.is_error:
        raise RuntimeError(f"GET {path} failed: {response.status_code} {response.text}")
    return response.json()


def post_json(base_url: str, path: str, payload: dict[str, Any]) -> dict[str, Any]:
    url = f"{base_url.rstrip('/')}{path}"
    response = httpx.post(url, json=payload, timeout=600.0)
    if response.is_error:
        raise RuntimeError(f"POST {path} failed: {response.status_code} {response.text}")
    return response.json()


def read_excel(file_path: Path) -> pd.DataFrame:
    return pd.read_excel(file_path)


def read_docx_preview(file_path: Path, max_paragraphs: int = 30) -> list[str]:
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


def split_law_basis(value: Any) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = text.replace(";", "；").split("；")
    return [part.strip() for part in parts if part.strip()]


def build_distribution(df: pd.DataFrame, column: str, top_n: int = 8) -> list[tuple[str, int, float]]:
    counter: Counter[str] = Counter()
    for value in df[column].fillna(""):
        text = str(value).strip()
        if text:
            counter[text] += 1
    total = sum(counter.values())
    if total == 0:
        return []
    return [
        (name, count, round(count * 100 / total, 1))
        for name, count in counter.most_common(top_n)
    ]


def build_law_distribution(df: pd.DataFrame, top_n: int = 20) -> list[tuple[str, int, float]]:
    counter: Counter[str] = Counter()
    case_count = max(len(df), 1)
    for value in df["法律依据（法律名称+条文）"].fillna(""):
        for item in split_law_basis(value):
            counter[item] += 1
    return [
        (name, count, round(count * 100 / case_count, 1))
        for name, count in counter.most_common(top_n)
    ]


def filter_cases(
    df: pd.DataFrame,
    selected_levels: list[str],
    selected_trials: list[str],
    selected_causes: list[str],
    keyword: str,
) -> pd.DataFrame:
    filtered = df.copy()
    if selected_levels:
        filtered = filtered[filtered["法院层级"].isin(selected_levels)]
    if selected_trials:
        filtered = filtered[filtered["法院审理程序"].isin(selected_trials)]
    if selected_causes:
        filtered = filtered[filtered["案由"].isin(selected_causes)]
    key = keyword.strip()
    if key:
        mask = (
            filtered["案号"].astype(str).str.contains(key, case=False, na=False)
            | filtered["审理法院"].astype(str).str.contains(key, case=False, na=False)
            | filtered["法院认为（精简后）"].astype(str).str.contains(key, case=False, na=False)
            | filtered["主要原因"].astype(str).str.contains(key, case=False, na=False)
        )
        filtered = filtered[mask]
    return filtered


def run_excel_workflow(base_url: str, query: str, target_case_count: int) -> None:
    payload = {"query": query.strip(), "target_case_count": int(target_case_count)}
    result = post_json(base_url, "/api/workflows/excel/run", payload)
    output_path = Path(result["output_path"])
    if not output_path.exists():
        raise RuntimeError(f"Excel output file not found: {output_path}")
    st.session_state["excel_result"] = result
    st.session_state["excel_df"] = read_excel(output_path)


def run_word_workflow(
    base_url: str,
    query: str,
    submitter: str,
    report_title: str,
    report_date: str,
) -> None:
    payload: dict[str, Any] = {
        "query": query.strip(),
        "submitter": submitter.strip(),
        "report_title": report_title.strip(),
    }
    if report_date.strip():
        payload["report_date"] = report_date.strip()
    result = post_json(base_url, "/api/workflows/word/run", payload)
    output_path = Path(result["output_path"])
    if not output_path.exists():
        raise RuntimeError(f"Word output file not found: {output_path}")
    st.session_state["word_result"] = result
    st.session_state["word_preview"] = read_docx_preview(output_path)


def render_outputs_panel() -> None:
    st.markdown('<div class="panel-title">结果文件</div>', unsafe_allow_html=True)
    excel_result = st.session_state.get("excel_result")
    word_result = st.session_state.get("word_result")

    if excel_result:
        excel_path = Path(excel_result["output_path"])
        st.success(f"Excel 已生成：{excel_path.name}")
        st.caption(f"案例数：{excel_result['row_count']} | 命中数：{excel_result['case_hit_count']}")
        if excel_path.exists():
            st.download_button(
                "下载 Excel",
                data=excel_path.read_bytes(),
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    if word_result:
        word_path = Path(word_result["output_path"])
        st.success(f"Word 已生成：{word_path.name}")
        st.caption(
            f"案例数：{word_result['row_count']} | 附件数："
            f"{word_result['attachment_case_count']}"
        )
        if word_path.exists():
            st.download_button(
                "下载 Word",
                data=word_path.read_bytes(),
                file_name=word_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )


def render_analysis_tab(df: pd.DataFrame | None) -> None:
    if df is None or df.empty:
        st.info("请先生成 Excel 结果后查看类案综述。")
        return

    total = len(df)
    court_count = df["审理法院"].nunique()
    cause_count = df["案由"].nunique()
    metric_cols = st.columns(3)
    metric_cols[0].metric("有效类案样本", total)
    metric_cols[1].metric("审理法院数量", court_count)
    metric_cols[2].metric("案由类型数", cause_count)

    left_col, right_col = st.columns(2, gap="large")
    with left_col:
        st.markdown('<div class="panel-title">分布概览</div>', unsafe_allow_html=True)
        for label, column in (("法院层级", "法院层级"), ("审理程序", "法院审理程序")):
            dist = build_distribution(df, column)
            st.markdown(f"**{label}**")
            if not dist:
                st.write("暂无数据")
            for name, count, ratio in dist:
                st.write(f"- {name}: {count} 件（{ratio}%）")

    with right_col:
        st.markdown('<div class="panel-title">类案核心裁判观点</div>', unsafe_allow_html=True)
        top_viewpoints = (
            df[["主要原因", "案号"]]
            .dropna()
            .drop_duplicates(subset=["案号"])
            .head(8)
        )
        if top_viewpoints.empty:
            st.write("暂无数据")
        for idx, row in top_viewpoints.reset_index(drop=True).iterrows():
            reason = str(row["主要原因"]).strip()
            case_no = str(row["案号"]).strip()
            st.markdown(
                f"- {idx + 1}. {reason}\n\n  `{case_no}`"
            )

    st.markdown('<div class="panel-title">类案检索结果分析</div>', unsafe_allow_html=True)
    law_dist = build_law_distribution(df, top_n=3)
    if law_dist:
        lines = [
            f"{name}：{count} 件（占样本 {ratio}%）"
            for name, count, ratio in law_dist
        ]
        st.write("；".join(lines))
    else:
        st.write("暂无可统计法律依据。")


def render_case_tab(df: pd.DataFrame | None) -> None:
    if df is None or df.empty:
        st.info("请先生成 Excel 结果后查看案例。")
        return

    left_col, right_col = st.columns([1, 3], gap="large")
    with left_col:
        st.markdown('<div class="panel-title">筛选条件</div>', unsafe_allow_html=True)
        levels = sorted(df["法院层级"].dropna().astype(str).unique().tolist())
        trials = sorted(df["法院审理程序"].dropna().astype(str).unique().tolist())
        causes = sorted(df["案由"].dropna().astype(str).unique().tolist())
        selected_levels = st.multiselect("法院级别", levels)
        selected_trials = st.multiselect("审理程序", trials)
        selected_causes = st.multiselect("案由", causes)
        keyword = st.text_input("关键词", placeholder="案号/法院/法院认为/主要原因")

    filtered_df = filter_cases(
        df=df,
        selected_levels=selected_levels,
        selected_trials=selected_trials,
        selected_causes=selected_causes,
        keyword=keyword,
    )

    with right_col:
        st.markdown('<div class="panel-title">案例列表</div>', unsafe_allow_html=True)
        st.caption(f"筛选后 {len(filtered_df)} 件，当前展示前 30 件。")
        if filtered_df.empty:
            st.warning("当前筛选条件下没有案例。")
            return

        download_buffer = BytesIO()
        filtered_df.to_excel(download_buffer, index=False)
        st.download_button(
            "下载当前筛选结果",
            data=download_buffer.getvalue(),
            file_name="filtered_cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        for _, row in filtered_df.head(30).iterrows():
            court = str(row.get("审理法院", "")).strip()
            case_no = str(row.get("案号", "")).strip()
            cause = str(row.get("案由", "")).strip()
            title = f"{case_no} | {court}"
            with st.expander(title, expanded=False):
                st.write(f"**案由**：{cause}")
                st.write(f"**审理程序**：{row.get('法院审理程序', '')}")
                st.write(f"**法院层级**：{row.get('法院层级', '')}")
                st.write("**法院认为（精简后）**")
                st.write(str(row.get("法院认为（精简后）", "")).strip())
                st.write("**裁判结果**")
                st.write(str(row.get("裁判结果", "")).strip())
                st.write("**主要原因**")
                st.write(str(row.get("主要原因", "")).strip())
                st.write("**法律依据（法律名称+条文）**")
                st.write(str(row.get("法律依据（法律名称+条文）", "")).strip())


def render_law_tab(df: pd.DataFrame | None) -> None:
    if df is None or df.empty:
        st.info("请先生成 Excel 结果后查看法律法规。")
        return

    st.markdown('<div class="panel-title">法律依据分布</div>', unsafe_allow_html=True)
    law_dist = build_law_distribution(df)
    if not law_dist:
        st.warning("Excel 数据中没有可解析的法律依据。")
        return

    law_df = pd.DataFrame(
        [
            {"法律依据": name, "命中次数": count, "占样本比例(%)": ratio}
            for name, count, ratio in law_dist
        ]
    )
    st.dataframe(law_df, use_container_width=True, hide_index=True)

    st.markdown('<div class="panel-title">前 10 条法律依据</div>', unsafe_allow_html=True)
    for name, count, ratio in law_dist[:10]:
        st.markdown(
            f'<div class="law-item"><strong>{name}</strong><br>'
            f'<span class="muted">{count} 件 / 占样本 {ratio}%</span></div>',
            unsafe_allow_html=True,
        )


inject_styles()
init_state()

st.markdown(
    """
<div class="header-card">
  <div class="app-title">⚖️ 类案智裁Agent</div>
  <div class="muted">基于已有工作流 API，一站式生成类案检索表格与报告</div>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.subheader("后端连接")
    backend_base_url = st.text_input("API Base URL", value="http://127.0.0.1:8000")
    if st.button("健康检查", use_container_width=True):
        try:
            health = get_json(backend_base_url, "/healthz")
            st.success(f"后端可用：{health}")
        except Exception as exc:  # noqa: BLE001
            st.error(f"健康检查失败：{exc}")

st.markdown('<div class="search-card">', unsafe_allow_html=True)
query = st.text_input(
    "检索问题",
    value="上下班途中交通事故工伤案例",
    placeholder="输入法律问题或检索要点，如案情描述、争议焦点、时间地区",
)
row_one = st.columns([1, 1, 1, 1], gap="small")
target_case_count = row_one[0].number_input("目标案例数量", min_value=1, max_value=200, value=50)
run_excel_clicked = row_one[1].button("生成类案检索表格", type="primary", use_container_width=True)
run_word_clicked = row_one[2].button("生成类案检索报告", use_container_width=True)
run_all_clicked = row_one[3].button("一键生成全部", use_container_width=True)
row_two = st.columns([1, 1, 1], gap="small")
submitter = row_two[0].text_input("提交人", value="xxxxxx")
report_title = row_two[1].text_input("报告标题", value="案涉争议问题检索报告")
report_date = row_two[2].text_input("报告日期（可选）", value="")
st.markdown("</div>", unsafe_allow_html=True)

if run_excel_clicked or run_all_clicked:
    if not query.strip():
        st.error("检索问题不能为空。")
    else:
        with st.spinner("正在生成 Excel，请稍候..."):
            try:
                run_excel_workflow(
                    base_url=backend_base_url,
                    query=query,
                    target_case_count=int(target_case_count),
                )
                st.success("Excel 工作流执行完成。")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Excel 工作流失败：{exc}")

if run_word_clicked or run_all_clicked:
    if not query.strip():
        st.error("检索问题不能为空。")
    else:
        with st.spinner("正在生成 Word 报告，请稍候..."):
            try:
                run_word_workflow(
                    base_url=backend_base_url,
                    query=query,
                    submitter=submitter,
                    report_title=report_title,
                    report_date=report_date,
                )
                st.success("Word 工作流执行完成。")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Word 工作流失败：{exc}")

preview_col, output_col = st.columns([3, 1], gap="large")

with output_col:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    render_outputs_panel()
    st.markdown("</div>", unsafe_allow_html=True)
    if st.session_state.get("word_preview"):
        st.markdown('<div class="panel-card">', unsafe_allow_html=True)
        st.markdown('<div class="panel-title">报告预览（前 30 段）</div>', unsafe_allow_html=True)
        st.text("\n\n".join(st.session_state["word_preview"]))
        st.markdown("</div>", unsafe_allow_html=True)

with preview_col:
    tab_analysis, tab_cases, tab_law = st.tabs(["类案综述", "案例", "法律法规"])
    excel_df = st.session_state.get("excel_df")
    with tab_analysis:
        render_analysis_tab(excel_df)
    with tab_cases:
        render_case_tab(excel_df)
    with tab_law:
        render_law_tab(excel_df)
