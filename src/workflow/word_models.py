from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class OpeningContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    报告标题: str = Field(..., min_length=1)
    开篇亮明观点: str = Field(..., min_length=1)
    争议焦点段: str = Field(..., min_length=1)
    类案检索结论段: str = Field(..., min_length=1)


class AttachmentCaseSelection(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    选中序号: list[int] = Field(..., min_length=1)


class WordReportContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    报告标题: str = Field(..., min_length=1)
    开篇亮明观点: str = Field(..., min_length=1)
    待决案件案情简述: str = Field(..., min_length=1)
    第一部分_类案基本事实概括: str = Field(..., min_length=1)
    第一部分_检索方法: str = Field(..., min_length=1)
    第一部分_检索情况: str = Field(..., min_length=1)
    第一部分_本案关联性: str = Field(..., min_length=1)
    第二部分_类案核心裁判要旨: str = Field(..., min_length=1)
    第二部分_观点总结列表: list[str] = Field(..., min_length=1)
    第四部分_相关法律法规原文: list[str] = Field(..., min_length=1)
    第五部分_结果分析: list[str] = Field(..., min_length=3, max_length=3)
    应当参照类案: list[str] = Field(..., min_length=1)
    可以参考类案: list[str] = Field(..., min_length=1)
