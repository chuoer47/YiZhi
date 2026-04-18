from __future__ import annotations

from pathlib import Path


QUERY = "上班途中车祸工伤案例"
TARGET_CASE_COUNT = 50
REWRITE_QUERY_COUNT = 3
PAGE_SIZE = 5
MAX_CONCURRENCY = 5
DOTENV_PATH = Path(".env")
WORD_SUBMITTER = "xxxxxx"

HEADERS = [
    "审理法院",
    "案号",
    "案由",
    "法院审理程序",
    "法院层级",
    "法院认为（精简后）",
    "裁判结果",
    "主要原因",
    "关键裁判结论-权利类",
    "关键裁判结论-金额类",
    "关键裁判结论-行为类",
    "法律依据（法律名称+条文）",
]
