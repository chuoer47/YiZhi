from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class OpeningContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    报告标题: str = Field(..., min_length=1)
    开篇亮明观点: str = Field(..., min_length=1)
    争议焦点段: str = Field(..., min_length=1)
    类案检索结论段: str = Field(..., min_length=1)


class SectionOneContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    第一部分标题: str = Field(..., min_length=1)
    第一部分正文: str = Field(..., min_length=1)


class SectionTwoContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    第二部分标题: str = Field(..., min_length=1)
    检索方法段: str = Field(..., min_length=1)
    类案检索情况段: str = Field(..., min_length=1)
    本案关联性段: str = Field(..., min_length=1)


class SectionThreeTableRow(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    中级人民法院案例: str = Field(..., min_length=1)
    案由: str = Field(..., min_length=1)
    裁判要点与理由: str = Field(..., min_length=1)


class SectionThreeViewpoint(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    观点标题: str = Field(..., min_length=1)
    观点总结: str = Field(..., min_length=1)
    表格标题: str = Field(..., min_length=1)
    裁判要点表格: list[SectionThreeTableRow] = Field(..., min_length=1)


class SectionThreeContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    第三部分标题: str = Field(..., min_length=1)
    第三部分引言段: str = Field(..., min_length=1)
    分类说明段: str = Field(..., min_length=1)
    观点列表: list[SectionThreeViewpoint] = Field(..., min_length=1)


class SectionFourLawItem(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    法规名称与条款: str = Field(..., min_length=1)
    条文原文: str = Field(..., min_length=1)
    条文要点列表: list[str] = Field(..., min_length=1)


class SectionFourContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    第四部分标题: str = Field(..., min_length=1)
    第四部分引言段: str = Field(..., min_length=1)
    法规条文列表: list[SectionFourLawItem] = Field(..., min_length=1)


class SectionFiveContent(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    第五部分标题: str = Field(..., min_length=1)
    分析结论列表: list[str] = Field(..., min_length=3, max_length=3)


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
