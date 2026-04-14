from __future__ import annotations

import sys
from pathlib import Path

from fastapi import FastAPI
from pydantic import BaseModel, Field

# Allow direct run: python src/workflow_api/app.py
SRC_ROOT = Path(__file__).resolve().parents[1]
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from workflow.excel_workflow import TARGET_CASE_COUNT, run_excel_workflow
from workflow.word_workflow import WORD_SUBMITTER, run_word_workflow

app = FastAPI(
    title="YiZhi Workflow API",
    version="1.0.0",
    description="Expose Excel and Word workflows as HTTP APIs.",
)


class ExcelRunRequest(BaseModel):
    query: str = Field(..., min_length=1, description="检索主题")
    target_case_count: int = Field(
        default=TARGET_CASE_COUNT,
        ge=1,
        description="目标类案数量",
    )


class ExcelRunResponse(BaseModel):
    output_path: str
    row_count: int
    case_hit_count: int


class WordRunRequest(BaseModel):
    query: str = Field(..., min_length=1, description="检索主题")
    submitter: str = Field(default=WORD_SUBMITTER, min_length=1, description="提交人")
    report_title: str = Field(default="案涉争议问题检索报告", min_length=1, description="报告标题")
    report_date: str | None = Field(default=None, description="报告日期，例如 2026年04月")


class WordRunResponse(BaseModel):
    output_path: str
    row_count: int
    attachment_case_count: int


@app.get("/healthz")
async def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/workflows/excel/run", response_model=ExcelRunResponse)
async def run_excel(request: ExcelRunRequest) -> ExcelRunResponse:
    excel_path, rows, case_hits = await run_excel_workflow(
        query=request.query,
        target_case_count=request.target_case_count,
    )
    return ExcelRunResponse(
        output_path=str(excel_path.resolve()),
        row_count=len(rows),
        case_hit_count=len(case_hits),
    )


@app.post("/api/workflows/word/run", response_model=WordRunResponse)
async def run_word(request: WordRunRequest) -> WordRunResponse:
    report_path, rows, attachment_case_hits = await run_word_workflow(
        query=request.query,
        submitter=request.submitter,
        report_title=request.report_title,
        report_date=request.report_date,
    )
    return WordRunResponse(
        output_path=str(report_path.resolve()),
        row_count=len(rows),
        attachment_case_count=len(attachment_case_hits),
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("workflow_api.app:app", host="127.0.0.1", port=8000, reload=False)
