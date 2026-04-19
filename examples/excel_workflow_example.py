# 用途：演示 Excel 工作流，生成类案检索表格。
# 运行命令：python examples/excel_workflow_example.py

from __future__ import annotations

import asyncio
import sys
from pathlib import Path

SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from workflow.excel_workflow import run_excel_workflow


QUERY = "上班途中车祸工伤案例"
TARGET_CASE_COUNT = 10


async def main() -> None:
    excel_path, rows, case_hits = await run_excel_workflow(
        query=QUERY,
        target_case_count=TARGET_CASE_COUNT,
    )
    print(f"Generated Excel: {excel_path.resolve()}")
    print(f"Rows: {len(rows)}")
    print(f"Case Hits: {len(case_hits)}")


if __name__ == "__main__":
    asyncio.run(main())
