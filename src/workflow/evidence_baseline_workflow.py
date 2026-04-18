from __future__ import annotations

import asyncio
import json
from datetime import datetime
from pathlib import Path

if __package__:
    from .case_pipeline import build_rows
    from .evidence_baseline_models import EvidenceBaselineData
    from .evidence_baseline_stats import (
        build_distribution,
        build_law_basis_stats,
    )
    from .workflow_config import QUERY
    from .workflow_runtime import load_env
else:
    from case_pipeline import build_rows
    from evidence_baseline_models import EvidenceBaselineData
    from evidence_baseline_stats import (
        build_distribution,
        build_law_basis_stats,
    )
    from workflow_config import QUERY
    from workflow_runtime import load_env


OUTPUT_DIR = Path("outputs")


def build_evidence_baseline(query: str, rows: list[object]) -> EvidenceBaselineData:
    total = len(rows)
    return EvidenceBaselineData(
        查询语句=query,
        样本总数=total,
        法院层级分布=build_distribution(rows, "法院层级"),
        审理程序分布=build_distribution(rows, "法院审理程序"),
        法律依据分布=build_law_basis_stats(rows, total),
    )


def save_evidence_baseline(data: EvidenceBaselineData) -> Path:
    output_path = (
        OUTPUT_DIR
        / f"evidence_baseline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    )
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
