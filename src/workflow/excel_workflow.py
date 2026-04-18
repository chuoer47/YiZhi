from __future__ import annotations

import asyncio
import sys
from pathlib import Path

from deli_api import CaseHit

# Allow direct run: python src/workflow/excel_workflow.py
SRC_ROOT = Path(__file__).resolve().parents[1]
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

if __package__:
    from .case_pipeline import (
        build_case_dedup_key,
        build_law_basis,
        build_rows,
        extract_case_fields,
        fetch_case_hits,
        process_hit,
        rewrite_law_queries,
        rewrite_queries,
    )
    from .excel_export import save_excel, save_excel_sync
    from .excel_models import (
        CaseExtractResult,
        KeyConclusions,
        LawQueryRewriteResult,
        MinimalCaseRow,
        QueryRewriteResult,
    )
    from .workflow_config import (
        DOTENV_PATH,
        HEADERS,
        MAX_CONCURRENCY,
        PAGE_SIZE,
        QUERY,
        REWRITE_QUERY_COUNT,
        TARGET_CASE_COUNT,
    )
    from .workflow_runtime import init_llm, load_env
else:
    from case_pipeline import (
        build_case_dedup_key,
        build_law_basis,
        build_rows,
        extract_case_fields,
        fetch_case_hits,
        process_hit,
        rewrite_law_queries,
        rewrite_queries,
    )
    from excel_export import save_excel, save_excel_sync
    from excel_models import (
        CaseExtractResult,
        KeyConclusions,
        LawQueryRewriteResult,
        MinimalCaseRow,
        QueryRewriteResult,
    )
    from workflow_config import (
        DOTENV_PATH,
        HEADERS,
        MAX_CONCURRENCY,
        PAGE_SIZE,
        QUERY,
        REWRITE_QUERY_COUNT,
        TARGET_CASE_COUNT,
    )
    from workflow_runtime import init_llm, load_env


async def run_excel_workflow(
    query: str = QUERY,
    target_case_count: int = TARGET_CASE_COUNT,
) -> tuple[Path, list[MinimalCaseRow], list[CaseHit]]:
    load_env()
    rows, case_hits = await build_rows(query=query, target_case_count=target_case_count)
    excel_path = await save_excel(rows)
    return excel_path, rows, case_hits


async def main() -> None:
    excel_path, rows, _ = await run_excel_workflow()
    print(f"Generated Excel: {excel_path.resolve()}")
    print(f"Rows: {len(rows)}")
    print("Columns:", " | ".join(HEADERS))


if __name__ == "__main__":
    asyncio.run(main())
