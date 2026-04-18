from __future__ import annotations

import asyncio
import html
import json
import re
from typing import Any

from deli_api import CaseHit, DeliLegalClient, LawHit, init_case_client
from langchain_core.messages import HumanMessage, SystemMessage

if __package__:
    from .excel_models import (
        CaseExtractResult,
        LawQueryRewriteResult,
        MinimalCaseRow,
        QueryRewriteResult,
    )
    from .prompt_loader import render_prompt
    from .workflow_config import (
        DOTENV_PATH,
        MAX_CONCURRENCY,
        PAGE_SIZE,
        QUERY,
        REWRITE_QUERY_COUNT,
        TARGET_CASE_COUNT,
    )
    from .workflow_runtime import init_llm
else:
    from excel_models import (
        CaseExtractResult,
        LawQueryRewriteResult,
        MinimalCaseRow,
        QueryRewriteResult,
    )
    from prompt_loader import render_prompt
    from workflow_config import (
        DOTENV_PATH,
        MAX_CONCURRENCY,
        PAGE_SIZE,
        QUERY,
        REWRITE_QUERY_COUNT,
        TARGET_CASE_COUNT,
    )
    from workflow_runtime import init_llm


def build_case_dedup_key(hit: CaseHit) -> str:
    case_no = (hit.case_no or "").replace("（", "(").replace("）", ")").replace(" ", "")
    if not case_no:
        raise ValueError(f"Case missing case_no: {hit.title}")
    return case_no


def build_law_basis(hits: list[LawHit]) -> str:
    parts: list[str] = []
    for hit in hits:
        law_name = re.sub(r"<[^>]+>", "", html.unescape(hit.law_name or "")).strip()
        article_no = re.sub(r"<[^>]+>", "", html.unescape(hit.article_no or "")).strip()
        if law_name and article_no:
            parts.append(f"{law_name}{article_no}")
        if len(parts) >= 2:
            break
    return "；".join(parts)


async def rewrite_queries(rewriter: Any, query: str) -> list[str]:
    if not query.strip():
        raise ValueError("QUERY must not be empty.")
    schema_text = json.dumps(QueryRewriteResult.model_json_schema(), ensure_ascii=False, indent=2)
    result: QueryRewriteResult = await rewriter.ainvoke(
        [
            SystemMessage(content=render_prompt("excel/rewrite_queries_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "excel/rewrite_queries_human.txt",
                    query=query,
                    rewrite_query_count=REWRITE_QUERY_COUNT,
                    schema_text=schema_text,
                )
            ),
        ]
    )
    rewritten = [item.strip() for item in result.改写查询]
    if len(set(rewritten)) != len(rewritten):
        raise ValueError("Rewritten queries contain duplicates.")
    if any(item == query for item in rewritten):
        raise ValueError("Rewritten queries contain the original query.")
    return rewritten


async def fetch_case_hits(case_client: Any, query: str, target_count: int) -> list[CaseHit]:
    if target_count <= 0:
        raise ValueError("TARGET_CASE_COUNT must be greater than 0.")
    if PAGE_SIZE < 1 or PAGE_SIZE > 5:
        raise ValueError("PAGE_SIZE must be between 1 and 5.")
    page_no = 1
    hits: list[CaseHit] = []
    seen: set[str] = set()
    while len(hits) < target_count:
        page_hits = await asyncio.to_thread(
            case_client.search_cases,
            query,
            page_no=page_no,
            page_size=PAGE_SIZE,
        )
        if not page_hits:
            break
        for hit in page_hits:
            key = build_case_dedup_key(hit)
            if key in seen:
                continue
            seen.add(key)
            hits.append(hit)
            if len(hits) >= target_count:
                break
        if len(page_hits) < PAGE_SIZE:
            break
        page_no += 1
    return hits


async def extract_case_fields(extractor: Any, hit: CaseHit) -> CaseExtractResult:
    if not hit.title.strip():
        raise ValueError("Case title is empty.")
    if not hit.content.strip():
        raise ValueError(f"Case content is empty: {hit.title}")
    schema_text = json.dumps(CaseExtractResult.model_json_schema(), ensure_ascii=False, indent=2)
    return await extractor.ainvoke(
        [
            SystemMessage(content=render_prompt("excel/extract_case_fields_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "excel/extract_case_fields_human.txt",
                    case_title=hit.title,
                    court_name=hit.court_name or "",
                    case_no=hit.case_no or "",
                    cause=hit.cause or "",
                    level_of_trial=hit.level_of_trial or "",
                    case_type=hit.case_type or "",
                    case_content=hit.content,
                    schema_text=schema_text,
                )
            ),
        ]
    )


async def rewrite_law_queries(
    law_rewriter: Any,
    hit: CaseHit,
    extracted: CaseExtractResult,
) -> list[str]:
    schema_text = json.dumps(LawQueryRewriteResult.model_json_schema(), ensure_ascii=False, indent=2)
    result: LawQueryRewriteResult = await law_rewriter.ainvoke(
        [
            SystemMessage(content=render_prompt("excel/rewrite_law_queries_system.txt")),
            HumanMessage(
                content=render_prompt(
                    "excel/rewrite_law_queries_human.txt",
                    case_title=hit.title,
                    cause=extracted.案由,
                    major_reason=extracted.主要原因,
                    judgment_result=extracted.裁判结果,
                    schema_text=schema_text,
                )
            ),
        ]
    )
    rewritten = [item.strip() for item in result.法律检索查询 if item.strip()]
    base_queries = [
        f"{hit.title} {extracted.主要原因}".strip(),
        extracted.案由.strip(),
        (hit.cause or "").strip(),
        extracted.主要原因.strip(),
    ]
    merged: list[str] = []
    seen: set[str] = set()
    for q in [*rewritten, *base_queries]:
        if q and q not in seen:
            seen.add(q)
            merged.append(q)
    return merged


async def process_hit(
    extractor: Any,
    law_rewriter: Any,
    law_client: DeliLegalClient,
    hit: CaseHit,
    semaphore: asyncio.Semaphore,
) -> MinimalCaseRow:
    async with semaphore:
        if not hit.court_name or not hit.case_no:
            raise ValueError(f"Case metadata missing court_name/case_no: {hit.title}")
        extracted = await extract_case_fields(extractor, hit)
        law_queries = await rewrite_law_queries(law_rewriter, hit, extracted)
        law_basis = ""
        for query in law_queries:
            q = query.strip()
            if not q:
                continue
            law_hits = await asyncio.to_thread(
                law_client.search_laws,
                q,
                page_size=3,
                field_name="semantic",
            )
            law_basis = build_law_basis(law_hits)
            if law_basis:
                break
        if not law_basis:
            law_basis = "未检索到明确法律条文"
        return MinimalCaseRow(
            审理法院=hit.court_name.strip(),
            案号=hit.case_no.strip(),
            案由=extracted.案由.strip(),
            法院审理程序=extracted.法院审理程序.strip(),
            法院层级=extracted.法院层级.strip(),
            法院认为_精简后=extracted.法院认为_精简后.strip(),
            裁判结果=extracted.裁判结果.strip(),
            主要原因=extracted.主要原因.strip(),
            关键裁判结论_权利类=extracted.关键裁判结论.权利类.strip(),
            关键裁判结论_金额类=extracted.关键裁判结论.金额类.strip(),
            关键裁判结论_行为类=extracted.关键裁判结论.行为类.strip(),
            法律依据=law_basis.strip(),
        )


async def build_rows(
    query: str = QUERY,
    target_case_count: int = TARGET_CASE_COUNT,
) -> tuple[list[MinimalCaseRow], list[CaseHit]]:
    if not query.strip():
        raise ValueError("query must not be empty.")
    if target_case_count <= 0:
        raise ValueError("target_case_count must be greater than 0.")

    with init_case_client() as case_client, DeliLegalClient.from_env(DOTENV_PATH) as law_client:
        llm = init_llm()
        rewriter = llm.with_structured_output(QueryRewriteResult, method="function_calling")
        law_rewriter = llm.with_structured_output(LawQueryRewriteResult, method="function_calling")
        extractor = llm.with_structured_output(CaseExtractResult, method="function_calling")

        rewritten_queries = await rewrite_queries(rewriter, query)
        all_queries = [query, *rewritten_queries]

        merged_hits: list[CaseHit] = []
        seen: set[str] = set()
        for search_query in all_queries:
            hits = await fetch_case_hits(case_client, search_query, target_case_count)
            for hit in hits:
                key = build_case_dedup_key(hit)
                if key in seen:
                    continue
                seen.add(key)
                merged_hits.append(hit)
                if len(merged_hits) >= target_case_count:
                    break
            if len(merged_hits) >= target_case_count:
                break

        semaphore = asyncio.Semaphore(MAX_CONCURRENCY)
        tasks = [process_hit(extractor, law_rewriter, law_client, hit, semaphore) for hit in merged_hits]
        rows = await asyncio.gather(*tasks)
        return rows, merged_hits
