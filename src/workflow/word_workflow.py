from __future__ import annotations

import asyncio
from pathlib import Path

from deli_api import CaseHit

if __package__:
    from .case_pipeline import build_rows
    from .excel_models import MinimalCaseRow
    from .word_export import save_word
    from .word_generation import (
        generate_opening_content,
        generate_section_five_content,
        generate_section_four_content,
        generate_section_one_content,
        generate_section_three_content,
        generate_section_two_content,
        generate_word_report_content,
        select_attachment_cases,
    )
    from .word_models import (
        AttachmentCaseSelection,
        OpeningContent,
        SectionFiveContent,
        SectionFourContent,
        SectionFourLawItem,
        SectionOneContent,
        SectionThreeContent,
        SectionThreeTableRow,
        SectionThreeViewpoint,
        SectionTwoContent,
        WordReportContent,
    )
    from .workflow_config import QUERY, WORD_SUBMITTER
    from .workflow_runtime import load_env
else:
    from case_pipeline import build_rows
    from excel_models import MinimalCaseRow
    from word_export import save_word
    from word_generation import (
        generate_opening_content,
        generate_section_five_content,
        generate_section_four_content,
        generate_section_one_content,
        generate_section_three_content,
        generate_section_two_content,
        generate_word_report_content,
        select_attachment_cases,
    )
    from word_models import (
        AttachmentCaseSelection,
        OpeningContent,
        SectionFiveContent,
        SectionFourContent,
        SectionFourLawItem,
        SectionOneContent,
        SectionThreeContent,
        SectionThreeTableRow,
        SectionThreeViewpoint,
        SectionTwoContent,
        WordReportContent,
    )
    from workflow_config import QUERY, WORD_SUBMITTER
    from workflow_runtime import load_env


async def main() -> None:
    word_report_path, rows, attachment_case_hits = await run_word_workflow()
    print(f"Generated Word Report: {word_report_path.resolve()}")
    print(f"Rows: {len(rows)}")
    print(f"Attachment Cases: {len(attachment_case_hits)}")


async def run_word_workflow(
    query: str = QUERY,
    submitter: str = WORD_SUBMITTER,
    report_title: str = "案涉争议问题检索报告",
    report_date: str | None = None,
) -> tuple[Path, list[MinimalCaseRow], list[CaseHit]]:
    if not query.strip():
        raise ValueError("query must not be empty.")

    load_env()
    rows, case_hits = await build_rows(query=query)
    report_content = await generate_word_report_content(query, rows)
    attachment_case_hits = await select_attachment_cases(query, rows, case_hits)
    word_report_path = await save_word(
        case_title=query,
        rows=rows,
        attachment_case_hits=attachment_case_hits,
        report_content=report_content,
        report_title=report_title,
        submitter=submitter,
        report_date=report_date,
    )
    return word_report_path, rows, attachment_case_hits


if __name__ == "__main__":
    asyncio.run(main())
