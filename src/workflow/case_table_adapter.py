from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field

if __package__:
    from .excel_workflow import MinimalCaseRow
else:
    from excel_workflow import MinimalCaseRow


class CaseTableRow(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    中级人民法院案例: str = Field(..., min_length=1)
    案由: str = Field(..., min_length=1)
    裁判要点与理由: str = Field(..., min_length=1)


def _compose_case_label(court_name: str, case_no: str) -> str:
    court = court_name.strip()
    case = case_no.strip()
    if court and case:
        return f"{court}\n{case}"
    if court:
        return court
    if case:
        return case
    raise ValueError("court_name and case_no cannot both be empty")


def build_case_table_rows(rows: list[MinimalCaseRow]) -> list[CaseTableRow]:
    if not rows:
        raise ValueError("rows must not be empty")

    table_rows: list[CaseTableRow] = []
    for row in rows:
        table_rows.append(
            CaseTableRow(
                中级人民法院案例=_compose_case_label(row.审理法院, row.案号),
                案由=row.案由.strip(),
                裁判要点与理由=row.主要原因.strip(),
            )
        )
    return table_rows


def _extract_keywords(text: str) -> list[str]:
    parts = re.split(r"[，。；：、\s（）()【】《》\[\]{}<>\-—,.;:!?\n\r\t]+", text)
    return [part for part in parts if len(part) >= 2]


def build_case_table_rows_for_viewpoint(
    rows: list[MinimalCaseRow],
    viewpoint_text: str,
    max_cases: int = 3,
) -> list[CaseTableRow]:
    if not rows:
        raise ValueError("rows must not be empty")
    if not viewpoint_text.strip():
        raise ValueError("viewpoint_text must not be empty")
    if max_cases <= 0:
        raise ValueError("max_cases must be greater than 0")

    keywords = _extract_keywords(viewpoint_text)
    scored_rows: list[tuple[int, MinimalCaseRow]] = []
    for row in rows:
        haystack = " ".join(
            [
                row.案由,
                row.主要原因,
                row.裁判结果,
                row.法院认为_精简后,
            ]
        )
        score = sum(1 for kw in keywords if kw in haystack)
        scored_rows.append((score, row))

    scored_rows.sort(key=lambda item: item[0], reverse=True)
    positive_rows = [row for score, row in scored_rows if score > 0][:max_cases]
    selected_rows = positive_rows if positive_rows else [row for _, row in scored_rows[:max_cases]]

    return build_case_table_rows(selected_rows)
