# 用途：演示后端两种 API（Excel 生成接口 + Word 生成接口）的调用。
# 运行命令：python examples/workflow_api_example.py

from __future__ import annotations

import json

import httpx


BASE_URL = "http://127.0.0.1:8000"
QUERY = "上班途中车祸工伤案例"


def main() -> None:
    with httpx.Client(timeout=1800.0, trust_env=False) as client:
        health_resp = client.get(f"{BASE_URL}/healthz")
        health_resp.raise_for_status()
        print("[healthz]")
        print(json.dumps(health_resp.json(), ensure_ascii=False, indent=2))

        excel_payload = {
            "query": QUERY,
            "target_case_count": 5,
        }
        excel_resp = client.post(f"{BASE_URL}/api/workflows/excel/run", json=excel_payload)
        excel_resp.raise_for_status()
        print("\n[/api/workflows/excel/run]")
        print(json.dumps(excel_resp.json(), ensure_ascii=False, indent=2))

        word_payload = {
            "query": QUERY,
            "submitter": "xxxxxx",
            "report_title": "案涉争议问题检索报告",
        }
        word_resp = client.post(f"{BASE_URL}/api/workflows/word/run", json=word_payload)
        word_resp.raise_for_status()
        print("\n[/api/workflows/word/run]")
        print(json.dumps(word_resp.json(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
