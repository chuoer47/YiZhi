from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field

if __package__:
    from .workflow_config import REWRITE_QUERY_COUNT
else:
    from workflow_config import REWRITE_QUERY_COUNT


class KeyConclusions(BaseModel):
    权利类: str = Field(..., min_length=1)
    金额类: str = Field(..., min_length=1)
    行为类: str = Field(..., min_length=1)


class CaseExtractResult(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    案由: str = Field(..., min_length=1)
    法院审理程序: str = Field(..., min_length=1)
    法院层级: str = Field(..., min_length=1)
    法院认为_精简后: str = Field(..., alias="法院认为（精简后）", min_length=1)
    裁判结果: str = Field(..., min_length=1)
    主要原因: str = Field(..., min_length=1)
    关键裁判结论: KeyConclusions = Field(...)


class QueryRewriteResult(BaseModel):
    改写查询: list[str] = Field(
        ...,
        min_length=REWRITE_QUERY_COUNT,
        max_length=REWRITE_QUERY_COUNT,
    )


class LawQueryRewriteResult(BaseModel):
    法律检索查询: list[str] = Field(
        ...,
        min_length=REWRITE_QUERY_COUNT,
        max_length=REWRITE_QUERY_COUNT,
    )


class MinimalCaseRow(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    审理法院: str = Field(..., min_length=1)
    案号: str = Field(..., min_length=1)
    案由: str = Field(..., min_length=1)
    法院审理程序: str = Field(..., min_length=1)
    法院层级: str = Field(..., min_length=1)
    法院认为_精简后: str = Field(..., alias="法院认为（精简后）", min_length=1)
    裁判结果: str = Field(..., min_length=1)
    主要原因: str = Field(..., min_length=1)
    关键裁判结论_权利类: str = Field(..., min_length=1)
    关键裁判结论_金额类: str = Field(..., min_length=1)
    关键裁判结论_行为类: str = Field(..., min_length=1)
    法律依据: str = Field(..., alias="法律依据（法律名称+条文）", min_length=1)

    def to_excel_row(self) -> list[str]:
        return [
            self.审理法院,
            self.案号,
            self.案由,
            self.法院审理程序,
            self.法院层级,
            self.法院认为_精简后,
            self.裁判结果,
            self.主要原因,
            self.关键裁判结论_权利类,
            self.关键裁判结论_金额类,
            self.关键裁判结论_行为类,
            self.法律依据,
        ]
