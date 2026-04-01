from __future__ import annotations

import json
from collections.abc import Mapping
from typing import Any

from langchain_core.tools import StructuredTool

from deli_api.client import DeliLegalClient
from deli_api.schemas import CaseSearchParams, LawSearchParams


def create_search_laws_tool(
    *,
    client: DeliLegalClient,
    defaults: Mapping[str, Any] | None = None,
    name: str = "search_laws",
    description: str | None = None,
) -> StructuredTool:
    default_params = dict(defaults or {})

    def _search_laws(
        query: str,
        page_no: int = 1,
        page_size: int = 5,
        sort_field: str = "correlation",
        sort_order: str = "desc",
        time_liness_type_arr: list[str] | None = None,
        publish_year_start: str | None = None,
        publish_year_end: str | None = None,
        active_year_start: str | None = None,
        active_year_end: str | None = None,
        field_name: str = "semantic",
    ) -> str:
        params = {
            "query": query,
            "page_no": page_no,
            "page_size": page_size,
            "sort_field": sort_field,
            "sort_order": sort_order,
            "time_liness_type_arr": time_liness_type_arr,
            "publish_year_start": publish_year_start,
            "publish_year_end": publish_year_end,
            "active_year_start": active_year_start,
            "active_year_end": active_year_end,
            "field_name": field_name,
        }
        hits = client.search_laws(**_merge_defaults(default_params, params))
        return _serialize_hits(hits)

    return StructuredTool.from_function(
        func=_search_laws,
        name=name,
        description=description
        or "Search laws, regulations, and judicial interpretations from DeliLegal.",
        args_schema=LawSearchParams,
    )


def create_search_cases_tool(
    *,
    client: DeliLegalClient,
    defaults: Mapping[str, Any] | None = None,
    name: str = "search_cases",
    description: str | None = None,
) -> StructuredTool:
    default_params = dict(defaults or {})

    def _search_cases(
        query: str,
        page_no: int = 1,
        page_size: int = 5,
        sort_field: str = "correlation",
        sort_order: str = "desc",
        case_year_start: str | None = None,
        case_year_end: str | None = None,
        court_level_arr: list[str] | None = None,
    ) -> str:
        params = {
            "query": query,
            "page_no": page_no,
            "page_size": page_size,
            "sort_field": sort_field,
            "sort_order": sort_order,
            "case_year_start": case_year_start,
            "case_year_end": case_year_end,
            "court_level_arr": court_level_arr,
        }
        hits = client.search_cases(**_merge_defaults(default_params, params))
        return _serialize_hits(hits)

    return StructuredTool.from_function(
        func=_search_cases,
        name=name,
        description=description or "Search legal cases from DeliLegal.",
        args_schema=CaseSearchParams,
    )


def create_legal_search_tools(
    *,
    client: DeliLegalClient,
    law_defaults: Mapping[str, Any] | None = None,
    case_defaults: Mapping[str, Any] | None = None,
) -> list[StructuredTool]:
    return [
        create_search_laws_tool(client=client, defaults=law_defaults),
        create_search_cases_tool(client=client, defaults=case_defaults),
    ]


def _merge_defaults(defaults: Mapping[str, Any], params: Mapping[str, Any]) -> dict[str, Any]:
    merged = dict(defaults)
    for key, value in params.items():
        if value is not None:
            merged[key] = value
    return merged


def _serialize_hits(hits: list[Any]) -> str:
    payload = []
    for hit in hits:
        if hasattr(hit, "model_dump"):
            payload.append(hit.model_dump(exclude={"raw"}))
        else:
            payload.append(hit)
    return json.dumps(payload, ensure_ascii=False, indent=2)
