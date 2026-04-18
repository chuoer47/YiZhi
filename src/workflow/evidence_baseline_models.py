from __future__ import annotations

from pydantic import BaseModel, Field


class EvidenceStat(BaseModel):
    观点: str = Field(..., description="观点文本或分类名称")
    支持数量: int = Field(..., ge=0, description="支持该观点的样本数量")
    支持占比: float = Field(
        ...,
        ge=0,
        le=100,
        description="在全部样本中的占比，百分比",
    )
    案号列表: list[str] = Field(
        default_factory=list,
        description="支持该观点的案号样本",
    )


class EvidenceBaselineData(BaseModel):
    查询语句: str
    样本总数: int
    法院层级分布: dict[str, int]
    审理程序分布: dict[str, int]
    法律依据分布: list[EvidenceStat]
