from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any

if __package__:
    from .evidence_baseline_models import EvidenceStat
else:
    from evidence_baseline_models import EvidenceStat


TOP_K = 12
UNKNOWN_VALUES = {"", "未明确", "未检索到明确法律条文", "不适用", "无"}


def normalize_value(value: str) -> str:
    return " ".join((value or "").split()).strip()


def build_distribution(rows: list[Any], field_name: str) -> dict[str, int]:
    counter: Counter[str] = Counter()
    for row in rows:
        value = normalize_value(getattr(row, field_name))
        if not value:
            continue
        counter[value] += 1
    return dict(counter.most_common())


def build_viewpoint_stats(
    rows: list[Any],
    field_name: str,
    total: int,
) -> list[EvidenceStat]:
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


def build_law_basis_stats(rows: list[Any], total: int) -> list[EvidenceStat]:
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
