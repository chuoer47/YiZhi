from __future__ import annotations

from typing import Any, Literal

from langchain_core.documents import Document
from pydantic import BaseModel, ConfigDict, Field


class CaseHit(BaseModel):
    source_type: Literal["case"] = "case"
    title: str
    content: str = ""
    score: float | None = None
    source_id: str | None = None
    citation: str | None = None
    url: str | None = None
    case_no: str | None = None
    court_name: str | None = None
    case_type: str | None = None
    cause: str | None = None
    case_date: str | None = None
    level_of_trial: str | None = None
    judgement_type: str | None = None
    publish_type: str | None = None
    publish_type_name: str | None = None
    raw: dict[str, Any] = Field(default_factory=dict)

    def to_document(self, include_raw: bool = False) -> Document:
        metadata = {
            "source_type": self.source_type,
            "title": self.title,
            "score": self.score,
            "source_id": self.source_id,
            "citation": self.citation,
            "url": self.url,
            "case_no": self.case_no,
            "court_name": self.court_name,
            "case_type": self.case_type,
            "cause": self.cause,
            "case_date": self.case_date,
            "level_of_trial": self.level_of_trial,
            "judgement_type": self.judgement_type,
            "publish_type": self.publish_type,
            "publish_type_name": self.publish_type_name,
        }
        metadata = {k: v for k, v in metadata.items() if v is not None}
        if include_raw:
            metadata["raw"] = self.raw
        return Document(page_content=self.content.strip() or self.title, metadata=metadata)


class LawHit(BaseModel):
    source_type: Literal["law"] = "law"
    title: str
    content: str = ""
    score: float | None = None
    source_id: str | None = None
    citation: str | None = None
    url: str | None = None
    law_name: str | None = None
    article_no: str | None = None
    publisher_name: str | None = None
    level_name: str | None = None
    publish_date: str | None = None
    active_date: str | None = None
    timeliness_type: str | None = None
    raw: dict[str, Any] = Field(default_factory=dict)

    def to_document(self, include_raw: bool = False) -> Document:
        metadata = {
            "source_type": self.source_type,
            "title": self.title,
            "score": self.score,
            "source_id": self.source_id,
            "citation": self.citation,
            "url": self.url,
            "law_name": self.law_name,
            "article_no": self.article_no,
            "publisher_name": self.publisher_name,
            "level_name": self.level_name,
            "publish_date": self.publish_date,
            "active_date": self.active_date,
            "timeliness_type": self.timeliness_type,
        }
        metadata = {k: v for k, v in metadata.items() if v is not None}
        if include_raw:
            metadata["raw"] = self.raw
        return Document(page_content=self.content.strip() or self.title, metadata=metadata)


class LegalHit(BaseModel):
    """Backward-compatible generic hit model."""

    source_type: Literal["law", "case"]
    title: str
    content: str = ""
    score: float | None = None
    source_id: str | None = None
    citation: str | None = None
    url: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    raw: dict[str, Any] = Field(default_factory=dict)

    def to_document(self, include_raw: bool = False) -> Document:
        metadata = {
            "source_type": self.source_type,
            "title": self.title,
            "score": self.score,
            "source_id": self.source_id,
            "citation": self.citation,
            "url": self.url,
            **self.metadata,
        }
        if include_raw:
            metadata["raw"] = self.raw
        return Document(page_content=self.content.strip() or self.title, metadata=metadata)


class LawSearchParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    query: str = Field(..., description="Search query for laws or regulations.")
    page_no: int = Field(1, ge=1, description="Page number, starting from 1.")
    page_size: int = Field(5, ge=1, le=50, description="Number of hits per page.")
    sort_field: str = Field("correlation", description="Sort field. Usually correlation.")
    sort_order: str = Field("desc", description="Sort order. Usually desc.")
    time_liness_type_arr: list[str] = Field(
        default_factory=lambda: ["5"],
        description="Timeliness filter used by the upstream law search API.",
    )
    publish_year_start: str | None = Field(
        None,
        description="Publish date start, format YYYY-MM-DD.",
    )
    publish_year_end: str | None = Field(
        None,
        description="Publish date end, format YYYY-MM-DD.",
    )
    active_year_start: str | None = Field(
        None,
        description="Active date start, format YYYY-MM-DD.",
    )
    active_year_end: str | None = Field(
        None,
        description="Active date end, format YYYY-MM-DD.",
    )
    field_name: str = Field(
        "semantic",
        description="Upstream search field name. The sample uses semantic.",
    )


class CaseSearchParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    query: str = Field(..., description="Search query for legal cases.")
    page_no: int = Field(1, ge=1, description="Page number, starting from 1.")
    page_size: int = Field(5, ge=1, le=50, description="Number of hits per page.")
    sort_field: str = Field("correlation", description="Sort field. Usually correlation.")
    sort_order: str = Field("desc", description="Sort order. Usually desc.")
    case_year_start: str | None = Field(
        None,
        description="Case date start, format YYYY-MM-DD.",
    )
    case_year_end: str | None = Field(
        None,
        description="Case date end, format YYYY-MM-DD.",
    )
    court_level_arr: list[str] = Field(
        default_factory=list,
        description="Upstream court level filter, for example ['0'].",
    )
