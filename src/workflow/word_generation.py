from __future__ import annotations

import json
from typing import Any

from deli_api import CaseHit
from langchain_core.messages import HumanMessage, SystemMessage

if __package__:
    from .evidence_baseline_workflow import build_evidence_baseline
    from .excel_models import MinimalCaseRow
    from .prompt_loader import render_prompt
    from .word_models import (
        AttachmentCaseSelection,
        OpeningContent,
        WordReportContent,
    )
    from .workflow_runtime import init_llm
else:
    from evidence_baseline_workflow import build_evidence_baseline
    from excel_models import MinimalCaseRow
    from prompt_loader import render_prompt
    from word_models import (
        AttachmentCaseSelection,
        OpeningContent,
        WordReportContent,
    )
    from workflow_runtime import init_llm


async def generate_opening_content(query: str, rows: list[MinimalCaseRow]) -> OpeningContent:
    if not query:
        raise ValueError("QUERY must not be empty.")
    if not rows:
        raise ValueError("rows must not be empty.")
    llm = init_llm()
    opening_writer = llm.with_structured_output(OpeningContent, method="function_calling")
    schema_text = json.dumps(OpeningContent.model_json_schema(), ensure_ascii=False, indent=2)
    rows_text = json.dumps([row.model_dump(by_alias=True) for row in rows[:12]], ensure_ascii=False, indent=2)
    return await opening_writer.ainvoke(
        [
            SystemMessage(content=render_prompt("word/writer_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "word/opening_human.txt",
                    query=query,
                    rows_text=rows_text,
                    schema_text=schema_text,
                )
            ),
        ]
    )


async def generate_word_report_content(query: str, rows: list[MinimalCaseRow]) -> WordReportContent:
    if not query:
        raise ValueError("QUERY must not be empty.")
    if not rows:
        raise ValueError("rows must not be empty.")
    llm = init_llm()
    writer = llm.with_structured_output(WordReportContent, method="function_calling")
    schema_text = json.dumps(WordReportContent.model_json_schema(), ensure_ascii=False, indent=2)
    rows_text = json.dumps([row.model_dump(by_alias=True) for row in rows], ensure_ascii=False, indent=2)
    evidence_baseline = build_evidence_baseline(query, rows)
    evidence_baseline_text = json.dumps(evidence_baseline.model_dump(), ensure_ascii=False, indent=2)
    return await writer.ainvoke(
        [
            SystemMessage(content=render_prompt("word/writer_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "word/report_human.txt",
                    query=query,
                    rows_text=rows_text,
                    evidence_baseline_text=evidence_baseline_text,
                    schema_text=schema_text,
                )
            ),
        ]
    )


async def select_attachment_cases(
    query: str,
    rows: list[MinimalCaseRow],
    case_hits: list[CaseHit],
) -> list[CaseHit]:
    if not query:
        raise ValueError("QUERY must not be empty.")
    if not case_hits:
        raise ValueError("case_hits must not be empty.")
    if len(rows) != len(case_hits):
        raise ValueError("rows and case_hits length mismatch.")
    llm = init_llm()
    selector = llm.with_structured_output(AttachmentCaseSelection, method="function_calling")
    schema_text = json.dumps(AttachmentCaseSelection.model_json_schema(), ensure_ascii=False, indent=2)

    candidates: list[dict[str, Any]] = []
    for idx, (row, hit) in enumerate(zip(rows, case_hits), start=1):
        candidates.append(
            {
                "序号": idx,
                "审理法院": hit.court_name,
                "案号": hit.case_no,
                "标题": hit.title,
                "案由": row.案由,
                "裁判结果": row.裁判结果,
                "主要原因": row.主要原因,
            }
        )
    candidates_text = json.dumps(candidates, ensure_ascii=False, indent=2)

    result: AttachmentCaseSelection = await selector.ainvoke(
        [
            SystemMessage(content=render_prompt("word/selector_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "word/attachment_selector_human.txt",
                    query=query,
                    candidates_text=candidates_text,
                    schema_text=schema_text,
                )
            ),
        ]
    )

    selected_hits: list[CaseHit] = []
    seen: set[int] = set()
    for i in result.选中序号:
        if i < 1 or i > len(case_hits):
            raise ValueError(f"Attachment case index out of range: {i}")
        if i in seen:
            continue
        seen.add(i)
        selected_hits.append(case_hits[i - 1])
    if not selected_hits:
        raise ValueError("Selected attachment case count must be greater than 0.")
    return selected_hits
