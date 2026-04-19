# 用途：演示 Word 工作流，生成类案检索报告文档。
# 运行命令：python examples/word_workflow_example.py

from __future__ import annotations

import asyncio
import sys
from pathlib import Path

SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from workflow.word_workflow import run_word_workflow


QUERY = "上班途中车祸工伤案例"
SUBMITTER = "xxxxxx"
REPORT_TITLE = "案涉争议问题检索报告"
REPORT_DATE = None


async def main() -> None:
    report_path, rows, attachment_case_hits = await run_word_workflow(
        query=QUERY,
        submitter=SUBMITTER,
        report_title=REPORT_TITLE,
        report_date=REPORT_DATE,
    )
    print(f"Generated Word Report: {report_path.resolve()}")
    print(f"Rows: {len(rows)}")
    print(f"Attachment Cases: {len(attachment_case_hits)}")


if __name__ == "__main__":
    asyncio.run(main())
