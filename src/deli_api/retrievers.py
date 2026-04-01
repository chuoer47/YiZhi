from __future__ import annotations

from langchain_core.documents import Document
from langchain_core.retrievers import BaseRetriever
from pydantic import Field

from deli_api.client import DeliLegalClient


class LawRetriever(BaseRetriever):
    client: DeliLegalClient
    page_size: int = 5
    page_no: int = 1
    sort_field: str = "correlation"
    sort_order: str = "desc"
    time_liness_type_arr: list[str] = Field(default_factory=lambda: ["5"])
    publish_year_start: str | None = None
    publish_year_end: str | None = None
    active_year_start: str | None = None
    active_year_end: str | None = None
    field_name: str = "semantic"

    def _get_relevant_documents(self, query: str, *, run_manager) -> list[Document]:
        hits = self.client.search_laws(
            query=query,
            page_no=self.page_no,
            page_size=self.page_size,
            sort_field=self.sort_field,
            sort_order=self.sort_order,
            time_liness_type_arr=self.time_liness_type_arr,
            publish_year_start=self.publish_year_start,
            publish_year_end=self.publish_year_end,
            active_year_start=self.active_year_start,
            active_year_end=self.active_year_end,
            field_name=self.field_name,
        )
        return [hit.to_document() for hit in hits]


class CaseRetriever(BaseRetriever):
    client: DeliLegalClient
    page_size: int = 5
    page_no: int = 1
    sort_field: str = "correlation"
    sort_order: str = "desc"
    case_year_start: str | None = None
    case_year_end: str | None = None
    court_level_arr: list[str] = Field(default_factory=list)

    def _get_relevant_documents(self, query: str, *, run_manager) -> list[Document]:
        hits = self.client.search_cases(
            query=query,
            page_no=self.page_no,
            page_size=self.page_size,
            sort_field=self.sort_field,
            sort_order=self.sort_order,
            case_year_start=self.case_year_start,
            case_year_end=self.case_year_end,
            court_level_arr=self.court_level_arr,
        )
        return [hit.to_document() for hit in hits]
