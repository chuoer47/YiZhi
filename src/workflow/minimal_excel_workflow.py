from __future__ import annotations

import asyncio
import html
import os
import re
import sys
import warnings
from pathlib import Path
from typing import Any

# Allow direct run: python src/workflow/minimal_excel_workflow.py
SRC_ROOT = Path(__file__).resolve().parents[1]
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from deli_api import CaseHit, DeliLegalClient, LawHit, init_case_client
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from pydantic import BaseModel, ConfigDict, Field


# =========================
# Debug Config (edit here)
# =========================
QUERY = "上班途中车祸工伤案例"
TARGET_CASE_COUNT = 50
REWRITE_QUERY_COUNT = 3
PAGE_SIZE = 5  # Deli API 每页最多 5 条
MAX_CONCURRENCY = 5
OUTPUT_XLSX = Path("outputs/minimal_case_table.xlsx")
DOTENV_PATH = Path(r"E:\\腾讯开悟-法律ai\\YiZhi\\.env")
DEBUG_LLM_RAW = False


HEADERS = [
    "审理法院",
    "案号",
    "案由",
    "法院审理程序",
    "法院层级",
    "法院认为（精简后）",
    "裁判结果",
    "主要原因",
    "关键裁判结论-权利类",
    "关键裁判结论-金额类",
    "关键裁判结论-行为类",
    "法律依据（法律名称+条文）",
]


class KeyConclusions(BaseModel):
    权利类: str = Field(default="未明确", description="权利义务方向的结论")
    金额类: str = Field(default="未明确", description="金额方向的结论")
    行为类: str = Field(default="未明确", description="行为义务方向的结论")


class CaseExtractResult(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    裁判结果: str = Field(description="判决主文/裁判结果")
    主要原因: str = Field(description="主要裁判原因，1-3句")
    关键裁判结论: KeyConclusions = Field(description="按三类输出的关键裁判结论")


def load_env() -> None:
    dotenv_path = DOTENV_PATH if DOTENV_PATH.exists() else None
    load_dotenv(dotenv_path=dotenv_path, override=False)


def init_llm() -> ChatOpenAI:
    api_key = os.getenv("LLM_API_KEY")
    if not api_key:
        raise RuntimeError("Missing LLM_API_KEY in environment.")

    return ChatOpenAI(
        model=os.getenv("LLM_MODEL"),
        api_key=api_key,
        temperature=0.1,
        base_url=os.getenv("LLM_BASE_URL"),
    )


def build_llm_extractor(llm: ChatOpenAI):
    return llm


def _suppress_known_pydantic_warning() -> None:
    warnings.filterwarnings(
        "ignore",
        message=r"^Pydantic serializer warnings:",
        category=UserWarning,
        module=r"pydantic\.main",
    )


async def extract_with_llm(
    extractor: Any,          
    case_title: str,
    raw_case_text: str,
) -> tuple[str, str, str, str, str]:
    prompt = f"""
你是专业的法律助理。请严格根据裁判文书原文提取以下字段，**必须返回合法的 JSON 对象**，不要添加任何解释、代码块或多余文字。
案件标题: {case_title}
裁判文书原文:
{raw_case_text}
请严格按以下 JSON Schema 输出（字段名必须完全一致）：
{{
  "裁判结果": "判决主文/裁判结果原文或精简版",
  "主要原因": "主要裁判原因，1-3句话",
  "关键裁判结论": {{
    "权利类": "权利义务方向的结论或'未明确'",
    "金额类": "金额方向的结论或'未明确'",
    "行为类": "行为义务方向的结论或'未明确'"
  }}
}}
只返回 JSON，不要加 ```json 标记。
    """.strip()
    _suppress_known_pydantic_warning()
    # 普通调用
    response = await extractor.ainvoke(prompt)
    raw_content = str(response.content or "").strip()
    if DEBUG_LLM_RAW:
        print("=== LLM Raw JSON Response ===")
        print(raw_content[:800] + "..." if len(raw_content) > 800 else raw_content)
    # 解析 JSON
    try:
        import json
        parsed = json.loads(raw_content)
        
        # 适配 Pydantic 模型（保持和原来完全一致的返回值）
        result = CaseExtractResult.model_validate(
            {
                "裁判结果": parsed.get("裁判结果", "未明确"),
                "主要原因": parsed.get("主要原因", "未明确"),
                "关键裁判结论": parsed.get("关键裁判结论", {}),
            }
        )
        
    except Exception as e:
        print(f"JSON 解析失败: {e}，使用兜底值")
        # 解析失败时兜底
        result = CaseExtractResult(
            裁判结果="未明确",
            主要原因="未明确",
            关键裁判结论=KeyConclusions()
        )
    judgment_result = result.裁判结果.strip() or "未明确"
    major_reason = result.主要原因.strip() or "未明确"
    key_obj = result.关键裁判结论
    key_right = key_obj.权利类.strip() or "未明确"
    key_amount = key_obj.金额类.strip() or "未明确"
    key_action = key_obj.行为类.strip() or "未明确"
    return judgment_result, major_reason, key_right, key_amount, key_action


async def process_hit(
    extractor: Any,
    law_client: DeliLegalClient,
    hit: CaseHit,
    semaphore: asyncio.Semaphore,
) -> list[str]:
    async with semaphore:
        court = str(hit.court_name or "")
        case_no = str(hit.case_no or hit.citation or "")
        cause = str(hit.cause or "未明确")
        trial_procedure = infer_trial_procedure(hit, case_no)
        court_level = infer_court_level(hit, court)
        raw_case_text = str(hit.content or "")

        (
            judgment_result,
            major_reason,
            key_right,
            key_amount,
            key_action,
        ) = await extract_with_llm(
            extractor=extractor,
            case_title=hit.title,
            raw_case_text=raw_case_text,
        )

        # LLM orders original reasoning into numbered points with keyword constraints;
        # helper falls back to rule-based ordering if output is invalid.
        court_reasoning_short = await build_ordered_reasoning_with_llm(extractor, raw_case_text)

        law_basis = await find_law_basis(
            law_client=law_client,
            case_title=hit.title,
            major_reason=major_reason,
            case_text=raw_case_text,
        )

        main_row = [
            court,
            case_no,
            cause,
            trial_procedure,
            court_level,
            court_reasoning_short,
            judgment_result,
            major_reason,
            key_right,
            key_amount,
            key_action,
            law_basis,
        ]
        return main_row


def infer_trial_procedure(hit: CaseHit, case_no: str) -> str:
    level = str(hit.level_of_trial or "").strip()
    if level:
        return level

    case_type = str(hit.case_type or "").strip()
    if "一审" in case_type:
        return "一审"
    if "二审" in case_type:
        return "二审"
    if "再审" in case_type:
        return "再审"

    normalized = _normalize_case_no(case_no)
    if "初" in normalized:
        return "一审"
    if "终" in normalized:
        return "二审"
    if "再" in normalized:
        return "再审"
    return "未明确"


def infer_court_level(hit: CaseHit, court_name: str) -> str:
    raw = "".join(
        [
            str(court_name or ""),
            str(hit.case_type or ""),
            str(hit.level_of_trial or ""),
        ]
    )
    if "最高" in raw:
        return "最高法院"
    if "高级" in raw or "高院" in raw:
        return "高院"
    if "中级" in raw or "中院" in raw:
        return "中院"
    if any(token in raw for token in ("基层", "人民法庭", "区人民法院", "县人民法院")):
        return "基层法院"
    if "人民法院" in raw:
        return "基层法院"
    return "未明确"


def extract_full_court_reasoning(raw_case_text: str) -> str:
    text = _clean_reasoning_text(raw_case_text)
    if not text:
        return "未明确"

    start_markers = ["争议焦点", "法院认为", "本院认为", "审理认为"]
    end_markers = ["判决如下", "裁定如下", "综上", "依照", "据此"]

    # Prefer "争议焦点" if present; otherwise prefer the last reasoning marker,
    # which is usually the substantive legal reasoning block rather than evidence review.
    start_idx = text.find("争议焦点")
    if start_idx == -1:
        candidates: list[int] = []
        for marker in start_markers:
            pos = text.find(marker)
            while pos != -1:
                candidates.append(pos)
                pos = text.find(marker, pos + 1)
        start_idx = max(candidates) if candidates else -1

    if start_idx == -1:
        return text[:2000]

    end_idx = len(text)
    for marker in end_markers:
        idx = text.find(marker, start_idx + 2)
        if idx != -1 and idx < end_idx:
            end_idx = idx

    reasoning = text[start_idx:end_idx].strip()
    return reasoning or text[:2000]


def _clean_reasoning_text(text: str | None) -> str:
    cleaned = str(text or "")
    # Remove upstream highlight tags like <em>...</em> and any other HTML tags.
    cleaned = re.sub(r"<[^>]+>", "", cleaned)
    cleaned = html.unescape(cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def _split_reasoning_parts(reasoning_text: str) -> list[str]:
    text = re.sub(r"\s+", " ", reasoning_text).strip()
    if not text:
        return []

    marker_pattern = r"(?=(?:^|\s)(?:[（(]?[一二三四五六七八九十]+[)）]|[一二三四五六七八九十]+、))"
    chunks = [chunk.strip(" ，。；;\n") for chunk in re.split(marker_pattern, text) if chunk.strip()]
    if len(chunks) >= 2:
        return chunks

    return [seg.strip(" ，。；;\n") for seg in re.split(r"[。；;]", text) if seg.strip()]


def _filter_reasoning_parts(parts: list[str]) -> list[str]:
    noise_tokens = (
        "本院依法认定如下事实",
        "经审理查明",
        "认定如下事实",
        "上述事实",
        "有以下证据",
        "证据如下",
        "经质证",
        "予以认定",
    )
    evidence_tokens = (
        "证据",
        "采信",
        "不予采信",
        "证据效力",
        "质证意见",
        "证明目的",
        "证明力",
        "事实如下",
    )
    focus_tokens = (
        "争议焦点",
        "本院认为",
        "法院认为",
        "是否",
        "应当",
        "不应",
        "符合",
        "不符合",
        "上下班途中",
        "工伤",
        "据此",
    )

    kept: list[str] = []
    for part in parts:
        compact = part.replace(" ", "")
        if any(token in compact for token in noise_tokens):
            continue
        if any(token in compact for token in evidence_tokens):
            continue
        if any(token in compact for token in focus_tokens) or len(compact) >= 20:
            kept.append(part)

    return kept


def _normalize_for_dedup(text: str) -> str:
    normalized = re.sub(r"[\s，。；：:、,.!?！？()（）\[\]【】'\"《》<>-]+", "", text)
    # Remove numbering prefixes like "1." / "一、" / "（一）"
    normalized = re.sub(r"^(?:\d+\.|[一二三四五六七八九十]+、|[（(][一二三四五六七八九十]+[)）])", "", normalized)
    return normalized


def _deduplicate_parts(parts: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for part in parts:
        key = _normalize_for_dedup(part)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(part)
    return result


def _strip_number_prefix(text: str) -> str:
    return re.sub(r"^\s*(?:\d+[\.、]|[（(]?\d+[)）]|[一二三四五六七八九十]+[、\.])\s*", "", text).strip()


def _format_numbered_points(parts: list[str], max_points: int = 4) -> str:
    selected = _deduplicate_parts(parts)[:max_points]
    if not selected:
        return "未明确"
    return "\n".join(f"{idx}. {item}" for idx, item in enumerate(selected, start=1))


def build_ordered_reasoning_from_raw(raw_case_text: str, max_points: int = 4) -> str:
    reasoning = extract_full_court_reasoning(raw_case_text)
    parts = _split_reasoning_parts(reasoning)
    filtered = _filter_reasoning_parts(parts)
    if not filtered:
        filtered = parts
    filtered = _deduplicate_parts(filtered)
    if not filtered:
        return "未明确"

    return _format_numbered_points(filtered, max_points=max_points)


async def build_ordered_reasoning_with_llm(
    llm: Any,
    raw_case_text: str,
    max_points: int = 4,
) -> str:
    reasoning = extract_full_court_reasoning(raw_case_text)
    if reasoning == "未明确":
        return reasoning

    prompt = f"""
你是法律文书整理助手。请仅对给定原文做“排序与筛选”，不要改写、不要总结、不要新增信息。

要求：
1. 仅保留【争议焦点】和【法院认为/本院认为】相关说理。
2. 必须排除包含以下表述的内容：
   - 事实认定如下
   - 本院依法认定如下事实
3. 输出最多{max_points}点，按律师阅读顺序排列。
4. 每行编号格式为：1. 2. 3. 4.
5. 只输出编号结果，不输出任何解释。

原文：
{reasoning}
    """.strip()

    try:
        msg = await llm.ainvoke(prompt)
        text = str(getattr(msg, "content", "") or "").strip()
    except Exception:
        text = ""

    if not text:
        return build_ordered_reasoning_from_raw(raw_case_text, max_points=max_points)

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    candidates: list[str] = []
    for line in lines:
        item = _strip_number_prefix(line)
        if item:
            candidates.append(item)

    if not candidates:
        return build_ordered_reasoning_from_raw(raw_case_text, max_points=max_points)

    filtered = _filter_reasoning_parts(candidates)
    if not filtered:
        filtered = candidates

    return _format_numbered_points(filtered, max_points=max_points)


def _compact_text(text: str, limit: int = 80) -> str:
    normalized = re.sub(r"\s+", "", text or "")
    return normalized[:limit]


async def search_laws_for_case(
    law_client: DeliLegalClient,
    query: str,
    page_size: int = 3,
) -> list[LawHit]:
    return await asyncio.to_thread(
        law_client.search_laws,
        query,
        page_size=page_size,
        field_name="semantic",
    )


async def find_law_basis(
    law_client: DeliLegalClient,
    case_title: str,
    major_reason: str,
    case_text: str,
) -> str:
    queries = [
        _compact_text(f"{case_title} {major_reason}", 120),
        _compact_text(case_title, 80),
        _compact_text(major_reason, 80),
        _compact_text(case_text, 80),
    ]

    seen_queries: set[str] = set()
    for query in queries:
        if not query or query in seen_queries:
            continue
        seen_queries.add(query)
        hits = await search_laws_for_case(law_client=law_client, query=query, page_size=3)
        if not hits:
            continue

        parts: list[str] = []
        for hit in hits[:2]:
            law_name = (hit.law_name or hit.title or "").strip()
            article_no = (hit.article_no or "").strip()
            citation = (hit.citation or "").strip()

            if law_name and article_no:
                parts.append(f"{law_name}{article_no}")
            elif citation:
                parts.append(citation)
            elif law_name:
                parts.append(law_name)

        if parts:
            return "；".join(parts)

    return "未检索到明确法律条文"


def _case_key(hit: CaseHit) -> str:
    return str(hit.source_id or hit.case_no or hit.citation or hit.title)


def _normalize_case_no(case_no: str | None) -> str:
    text = str(case_no or "").strip()
    if not text:
        return ""
    # Normalize common punctuation/whitespace variations in case numbers.
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    return text


def _case_dedup_key(hit: CaseHit) -> str:
    normalized_case_no = _normalize_case_no(hit.case_no)
    if normalized_case_no:
        return f"case_no:{normalized_case_no}"
    return f"fallback:{_case_key(hit)}"


async def rewrite_queries(llm: ChatOpenAI, query: str) -> list[str]:
    prompt = f"""
你是法律检索助手。请将用户输入改写为更专业、规范的法律检索语句。
输出{REWRITE_QUERY_COUNT}条不同的改写结果，每行一条。
不要解释，不要编号，不要引号。

用户输入：{query}
    """.strip()
    msg = await llm.ainvoke(prompt)
    text = str(msg.content or "")
    lines = [line.strip().lstrip("-").strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    uniq: list[str] = []
    seen: set[str] = set()
    for line in lines:
        if line == query or line in seen:
            continue
        seen.add(line)
        uniq.append(line)
        if len(uniq) >= REWRITE_QUERY_COUNT:
            break
    return uniq


async def fetch_case_hits(case_client: Any, query: str, target_count: int) -> list[CaseHit]:
    if target_count <= 0:
        return []

    page_size = max(1, min(PAGE_SIZE, 5))
    page_no = 1
    hits: list[CaseHit] = []
    seen: set[str] = set()

    while len(hits) < target_count:
        page_hits = await asyncio.to_thread(
            case_client.search_cases,
            query,
            page_no=page_no,
            page_size=page_size,
        )
        if not page_hits:
            break

        for hit in page_hits:
            key = _case_dedup_key(hit)
            if key in seen:
                continue
            seen.add(key)
            hits.append(hit)
            if len(hits) >= target_count:
                break

        if len(page_hits) < page_size:
            break
        page_no += 1

    return hits[:target_count]


async def build_rows() -> list[list[str]]:
    with init_case_client() as case_client:
        law_client = DeliLegalClient.from_env(DOTENV_PATH)
        llm = init_llm()
        extractor = build_llm_extractor(llm)

        rewritten_queries = await rewrite_queries(llm, QUERY)
        all_queries = [QUERY, *rewritten_queries]

        merged_hits: list[CaseHit] = []
        seen: set[str] = set()
        for q in all_queries:
            hits = await fetch_case_hits(case_client, q, TARGET_CASE_COUNT)
            for hit in hits:
                key = _case_dedup_key(hit)
                if key not in seen:
                    seen.add(key)
                    merged_hits.append(hit)
                if len(merged_hits) >= TARGET_CASE_COUNT:
                    break
            if len(merged_hits) >= TARGET_CASE_COUNT:
                break

        semaphore = asyncio.Semaphore(MAX_CONCURRENCY)
        tasks = [process_hit(extractor, law_client, hit, semaphore) for hit in merged_hits]
        rows = await asyncio.gather(*tasks)
        law_client.close()
        return rows


def _save_excel_sync(rows: list[list[str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "类案检索结果"

    ws.append(
        [
            "审理法院",
            "案号",
            "案由",
            "法院审理程序",
            "法院层级",
            "法院认为（精简后）",
            "裁判结果",
            "主要原因",
            "关键裁判结论",
            "",
            "",
            "法律依据（法律名称+条文）",
        ]
    )
    ws.append(["", "", "", "", "", "", "", "", "权利类", "金额类", "行为类", ""])

    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")
    ws.merge_cells("D1:D2")
    ws.merge_cells("E1:E2")
    ws.merge_cells("F1:F2")
    ws.merge_cells("G1:G2")
    ws.merge_cells("H1:H2")
    ws.merge_cells("I1:K1")
    ws.merge_cells("L1:L2")

    for row in rows:
        ws.append(row)

    for header_row in (1, 2):
        for cell in ws[header_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [16, 24, 22, 14, 14, 40, 36, 30, 24, 24, 24, 42]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + idx)].width = width

    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=12):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)


async def save_excel(rows: list[list[str]]) -> None:
    await asyncio.to_thread(_save_excel_sync, rows)


async def main() -> None:
    load_env()
    rows = await build_rows()
    await save_excel(rows)
    print(f"Generated: {OUTPUT_XLSX.resolve()}")
    print(f"Rows: {len(rows)}")
    print("Columns:", " | ".join(HEADERS))


if __name__ == "__main__":
    asyncio.run(main())
