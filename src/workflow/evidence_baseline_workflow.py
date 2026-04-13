from __future__ import annotations

import asyncio
import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from pydantic import BaseModel, Field

from excel_workflow import QUERY, MinimalCaseRow, build_rows, load_env


TOP_K = 12
OUTPUT_DIR = Path("outputs")
UNKNOWN_VALUES = {"", "未明确", "未检索到明确法律条文", "不适用", "无"}


class EvidenceStat(BaseModel):
    观点: str = Field(..., description="观点文本或分类名称")
    支持数量: int = Field(..., ge=0, description="支持该观点的样本数量")
    支持占比: float = Field(..., ge=0, le=100, description="在全部样本中的占比，百分比")
    案号列表: list[str] = Field(default_factory=list, description="支持该观点的案号样本")


class EvidenceBaselineData(BaseModel):
    查询语句: str
    样本总数: int
    法院层级分布: dict[str, int]
    审理程序分布: dict[str, int]
    法律依据分布: list[EvidenceStat]


def normalize_value(value: str) -> str:
    return " ".join((value or "").split()).strip()


def build_distribution(rows: list[MinimalCaseRow], field_name: str) -> dict[str, int]:
    counter: Counter[str] = Counter()
    for row in rows:
        value = normalize_value(getattr(row, field_name))
        if not value:
            continue
        counter[value] += 1
    return dict(counter.most_common())


def build_viewpoint_stats(rows: list[MinimalCaseRow], field_name: str, total: int) -> list[EvidenceStat]:
    cases_by_viewpoint: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        viewpoint = normalize_value(getattr(row, field_name))
        case_no = normalize_value(row.案号)
        if not viewpoint or viewpoint in UNKNOWN_VALUES:
            continue
        if not case_no:
            continue
        cases_by_viewpoint[viewpoint].add(case_no)

    stats: list[EvidenceStat] = []
    for viewpoint, case_nos in cases_by_viewpoint.items():
        count = len(case_nos)
        ratio = round((count / total) * 100, 1) if total else 0.0
        stats.append(
            EvidenceStat(
                观点=viewpoint,
                支持数量=count,
                支持占比=ratio,
                案号列表=sorted(case_nos),
            )
        )
    stats.sort(key=lambda item: item.支持数量, reverse=True)
    return stats[:TOP_K]


def build_law_basis_stats(rows: list[MinimalCaseRow], total: int) -> list[EvidenceStat]:
    cases_by_law_basis: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        case_no = normalize_value(row.案号)
        law_basis = normalize_value(row.法律依据)
        if not case_no or not law_basis or law_basis in UNKNOWN_VALUES:
            continue
        for item in law_basis.split("；"):
            law_item = normalize_value(item)
            if not law_item or law_item in UNKNOWN_VALUES:
                continue
            cases_by_law_basis[law_item].add(case_no)

    stats: list[EvidenceStat] = []
    for law_item, case_nos in cases_by_law_basis.items():
        count = len(case_nos)
        ratio = round((count / total) * 100, 1) if total else 0.0
        stats.append(
            EvidenceStat(
                观点=law_item,
                支持数量=count,
                支持占比=ratio,
                案号列表=sorted(case_nos),
            )
        )
    stats.sort(key=lambda item: item.支持数量, reverse=True)
    return stats[:TOP_K]


def build_evidence_baseline(query: str, rows: list[MinimalCaseRow]) -> EvidenceBaselineData:
    total = len(rows)
    return EvidenceBaselineData(
        查询语句=query,
        样本总数=total,
        法院层级分布=build_distribution(rows, "法院层级"),
        审理程序分布=build_distribution(rows, "法院审理程序"),
        法律依据分布=build_law_basis_stats(rows, total),
    )


def save_evidence_baseline(data: EvidenceBaselineData) -> Path:
    output_path = OUTPUT_DIR / f"evidence_baseline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(data.model_dump(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path


async def main() -> None:
    load_env()
    rows, _ = await build_rows()
    baseline = build_evidence_baseline(QUERY, rows)
    output_path = save_evidence_baseline(baseline)
    print(f"Generated Evidence Baseline: {output_path.resolve()}")
    print(f"Samples: {baseline.样本总数}")
    print(f"Court Levels: {len(baseline.法院层级分布)}")
    print(f"Law Basis Items: {len(baseline.法律依据分布)}")


if __name__ == "__main__":
    asyncio.run(main())
