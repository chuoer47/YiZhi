# 用途：演示 deli_api 的两类检索（法规检索 + 案例检索）。
# 运行命令：python examples/deli_api_example.py

from __future__ import annotations

import sys
from pathlib import Path

SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from deli_api import DeliLegalClient


LAW_QUERY = "深圳市房地产相关的法律规定有哪些？"
CASE_QUERY = "上班途中车祸工伤案例"


def main() -> None:
    with DeliLegalClient.from_env(".env") as client:
        law_hits = client.search_laws(
            query=LAW_QUERY,
            page_size=3,
            field_name="semantic",
            time_liness_type_arr=["5"],
        )
        case_hits = client.search_cases(
            query=CASE_QUERY,
            page_size=3,
        )

    print(f"[法律法规] 命中数量: {len(law_hits)}")
    for hit in law_hits:
        print(f"- {hit.title}")

    print(f"\n[类案] 命中数量: {len(case_hits)}")
    for hit in case_hits:
        print(f"- {hit.title}")


if __name__ == "__main__":
    main()
