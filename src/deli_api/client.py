from __future__ import annotations

import json
import os
from collections.abc import Mapping, Sequence
from typing import Any

import httpx
from dotenv import load_dotenv

from deli_api.exceptions import (
    DeliLegalAuthenticationError,
    DeliLegalConfigError,
    DeliLegalHTTPStatusError,
    DeliLegalRequestError,
    DeliLegalResponseDecodeError,
    DeliLegalResponseFormatError,
    DeliLegalUpstreamError,
)
from deli_api.schemas import LegalHit


SUCCESS_CODES = {0, 200, "0", "200", "success", "SUCCESS"}
AUTH_ERROR_CODES = {401, 403, "401", "403"}
AUTH_ERROR_TOKENS = (
    "appid",
    "secret",
    "auth",
    "authentication",
    "unauthorized",
    "forbidden",
    "signature",
)


class DeliLegalClient:
    def __init__(
        self,
        app_id: str,
        secret: str,
        *,
        base_url: str = "https://openapi.delilegal.com",
        timeout: float = 30.0,
        client: httpx.Client | None = None,
    ) -> None:
        self.app_id = app_id
        self.secret = secret
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self._client = client or httpx.Client(
            base_url=self.base_url,
            timeout=self.timeout,
            headers={
                "Content-Type": "application/json",
                "appid": self.app_id,
                "secret": self.secret,
            },
        )

    def __enter__(self) -> "DeliLegalClient":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    @classmethod
    def from_env(
        cls,
        dotenv_path: str | os.PathLike[str] | None = None,
        *,
        override: bool = False,
    ) -> "DeliLegalClient":
        resolved_dotenv_path = (
            os.fspath(dotenv_path) if dotenv_path is not None else None
        )
        load_dotenv(dotenv_path=resolved_dotenv_path, override=override)
        app_id = os.getenv("DELILEGAL_APP_ID")
        secret = os.getenv("DELILEGAL_SECRET")
        base_url = os.getenv("DELILEGAL_BASE_URL", "https://openapi.delilegal.com")
        timeout = float(os.getenv("DELILEGAL_TIMEOUT", "30"))

        missing = [
            name
            for name, value in (
                ("DELILEGAL_APP_ID", app_id),
                ("DELILEGAL_SECRET", secret),
            )
            if not value
        ]
        if missing:
            raise DeliLegalConfigError(
                "Missing required DeliLegal environment variables.",
                details={"missing": missing},
            )

        return cls(app_id=app_id, secret=secret, base_url=base_url, timeout=timeout)

    def close(self) -> None:
        self._client.close()

    def search_laws(
        self,
        query: str,
        *,
        page_no: int = 1,
        page_size: int = 5,
        sort_field: str = "correlation",
        sort_order: str = "desc",
        time_liness_type_arr: Sequence[str] | None = None,
        publish_year_start: str | None = None,
        publish_year_end: str | None = None,
        active_year_start: str | None = None,
        active_year_end: str | None = None,
        field_name: str = "semantic",
        extra_condition: Mapping[str, Any] | None = None,
    ) -> list[LegalHit]:
        condition: dict[str, Any] = {
            "keywords": [query],
            "fieldName": field_name,
        }
        if time_liness_type_arr:
            condition["timeLinessTypeArr"] = list(time_liness_type_arr)
        condition.update(
            _drop_none(
                {
                    "publishYearStart": publish_year_start,
                    "publishYearEnd": publish_year_end,
                    "activeYearStart": active_year_start,
                    "activeYearEnd": active_year_end,
                }
            )
        )
        if extra_condition:
            condition.update(extra_condition)

        payload = {
            "pageNo": page_no,
            "pageSize": page_size,
            "sortField": sort_field,
            "sortOrder": sort_order,
            "condition": condition,
        }
        items = self._post("/api/qa/v3/search/queryListLaw", payload)
        return [self._normalize_law_hit(item) for item in items]

    def search_cases(
        self,
        query: str,
        *,
        page_no: int = 1,
        page_size: int = 5,
        sort_field: str = "correlation",
        sort_order: str = "desc",
        case_year_start: str | None = None,
        case_year_end: str | None = None,
        court_level_arr: Sequence[str] | None = None,
        extra_condition: Mapping[str, Any] | None = None,
    ) -> list[LegalHit]:
        condition: dict[str, Any] = {
            "keywordArr": [query],
        }
        if court_level_arr:
            condition["courtLevelArr"] = list(court_level_arr)
        condition.update(
            _drop_none(
                {
                    "caseYearStart": case_year_start,
                    "caseYearEnd": case_year_end,
                }
            )
        )
        if extra_condition:
            condition.update(extra_condition)

        payload = {
            "pageNo": page_no,
            "pageSize": page_size,
            "sortField": sort_field,
            "sortOrder": sort_order,
            "condition": condition,
        }
        items = self._post("/api/qa/v3/search/queryListCase", payload)
        return [self._normalize_case_hit(item) for item in items]

    def _post(self, path: str, payload: dict[str, Any]) -> list[dict[str, Any]]:
        response: httpx.Response | None = None
        try:
            response = self._client.post(path, json=payload)
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            response = exc.response
            raise DeliLegalHTTPStatusError(
                "Upstream returned a non-2xx HTTP status.",
                status_code=response.status_code if response is not None else None,
                path=path,
                payload=payload,
                details={"method": "POST"},
                response_text=_truncate_text(response.text if response is not None else None),
            ) from exc
        except httpx.RequestError as exc:
            raise DeliLegalRequestError(
                "Request to upstream failed.",
                path=path,
                payload=payload,
                details={"method": "POST", "reason": str(exc)},
            ) from exc

        try:
            data = response.json()
        except ValueError as exc:
            raise DeliLegalResponseDecodeError(
                "Upstream API did not return valid JSON.",
                path=path,
                payload=payload,
                response_text=_truncate_text(response.text),
            ) from exc

        self._raise_for_api_error(data, path=path, payload=payload)

        found, items = _find_records(data)
        if not found:
            raise DeliLegalResponseFormatError(
                "Unable to locate a result list in upstream payload.",
                path=path,
                payload=payload,
                details={"top_level_keys": list(data.keys()) if isinstance(data, dict) else []},
                response_text=_truncate_text(json.dumps(data, ensure_ascii=False)),
            )

        return [item for item in items if isinstance(item, dict)]

    def _raise_for_api_error(
        self,
        response_payload: Any,
        *,
        path: str,
        payload: dict[str, Any],
    ) -> None:
        if not isinstance(response_payload, dict):
            return

        success = response_payload.get("success")
        code = response_payload.get("code")
        message = str(
            response_payload.get("message")
            or response_payload.get("msg")
            or "Unknown upstream error."
        )

        if success is False or (code is not None and code not in SUCCESS_CODES):
            error_cls = (
                DeliLegalAuthenticationError
                if _is_auth_error(code=code, message=message)
                else DeliLegalUpstreamError
            )
            raise error_cls(
                message,
                code=code,
                path=path,
                payload=payload,
                details=_drop_none(
                    {
                        "success": success,
                        "query_id": _extract_query_id(response_payload),
                    }
                ),
                response_text=_truncate_text(json.dumps(response_payload, ensure_ascii=False)),
            )

    def _normalize_law_hit(self, item: dict[str, Any]) -> LegalHit:
        law_name = _first_str(item, "lawName", "title", "name", "documentName")
        article_no = _first_str(item, "articleNo", "articleNum", "clauseNo", "article")
        title = law_name or _first_str(item, "title", "name") or "Untitled law"
        content = _first_str(
            item,
            "content",
            "articleContent",
            "summary",
            "snippet",
            "text",
            "abstract",
        )
        citation = _build_law_citation(law_name, article_no) or title
        metadata = _drop_none(
            {
                "law_name": law_name,
                "article_no": article_no,
                "publisher_name": _first_str(item, "publisherName", "publisher"),
                "level_name": _first_str(item, "levelName", "level"),
                "publish_date": _first_str(item, "publishDate", "publishTime", "publishYear"),
                "active_date": _first_str(item, "activeDate", "activeTime", "effectiveDate"),
                "timeliness_type": _first_str(
                    item,
                    "timelinessName",
                    "timeLinessType",
                    "timelinessType",
                ),
            }
        )
        return LegalHit(
            source_type="law",
            title=title,
            content=content or title,
            score=_first_float(item, "score", "correlation", "similarity"),
            source_id=_first_str(item, "id", "lawId", "documentId", "docId"),
            citation=citation,
            url=_first_str(item, "url", "link", "detailUrl"),
            metadata=metadata,
            raw=item,
        )

    def _normalize_case_hit(self, item: dict[str, Any]) -> LegalHit:
        case_no = _first_str(item, "caseNo", "caseNumber", "\u6848\u53f7")
        title = _first_str(item, "title", "caseTitle", "name") or case_no or "Untitled case"
        content = _first_str(
            item,
            "content",
            "summary",
            "snippet",
            "text",
            "judgmentReason",
            "abstract",
        )
        court_name = _first_str(item, "courtName", "court", "courtFullName")
        citation = case_no or title
        metadata = _drop_none(
            {
                "case_no": case_no,
                "court_name": court_name,
                "case_type": _first_str(item, "caseType", "trialProcedure"),
                "cause": _first_str(item, "cause", "caseCause"),
                "case_date": _first_str(
                    item,
                    "caseDate",
                    "judgementDate",
                    "judgmentDate",
                    "refereeDate",
                ),
                "level_of_trial": _first_str(item, "levelOfTrial"),
                "judgement_type": _first_str(item, "judgementType", "judgmentType"),
                "province": _first_str(item, "province"),
            }
        )
        return LegalHit(
            source_type="case",
            title=title,
            content=content or title,
            score=_first_float(item, "score", "correlation", "similarity"),
            source_id=_first_str(item, "id", "caseId", "documentId", "docId"),
            citation=citation,
            url=_first_str(item, "url", "link", "detailUrl"),
            metadata=metadata,
            raw=item,
        )


def init_case_client(
    dotenv_path: str | os.PathLike[str] | None = None,
    *,
    override: bool = False,
) -> DeliLegalClient:
    """One-line factory for creating a DeliLegal client from env vars."""
    return DeliLegalClient.from_env(dotenv_path=dotenv_path, override=override)


def _find_records(payload: Any) -> tuple[bool, list[Any]]:
    if isinstance(payload, list):
        return True, payload
    if not isinstance(payload, dict):
        return False, []

    for key in ("list", "records", "rows", "items", "result", "results", "data", "body"):
        if key in payload:
            found, result = _find_records(payload[key])
            if found:
                return True, result

    for value in payload.values():
        if isinstance(value, list) and all(isinstance(item, dict) for item in value):
            return True, value

    return False, []


def _first_str(payload: Mapping[str, Any], *keys: str) -> str | None:
    value = _first_value(payload, *keys)
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _first_float(payload: Mapping[str, Any], *keys: str) -> float | None:
    value = _first_value(payload, *keys)
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _first_value(payload: Mapping[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in payload and payload[key] not in (None, ""):
            return payload[key]
    return None


def _build_law_citation(law_name: str | None, article_no: str | None) -> str | None:
    if law_name and article_no:
        return f"{law_name}{article_no}"
    return law_name or article_no


def _extract_query_id(payload: Mapping[str, Any]) -> str | None:
    for key in ("queryId", "query_id"):
        value = _first_value(payload, key)
        if value:
            return str(value)
    body = payload.get("body")
    if isinstance(body, Mapping):
        return _extract_query_id(body)
    return None


def _is_auth_error(*, code: Any, message: str) -> bool:
    if code in AUTH_ERROR_CODES:
        return True
    normalized = message.lower()
    return any(token in normalized for token in AUTH_ERROR_TOKENS)


def _truncate_text(text: str | None, limit: int = 2000) -> str | None:
    if text is None:
        return None
    if len(text) <= limit:
        return text
    return f"{text[:limit]}...<truncated>"


def _drop_none(payload: Mapping[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in payload.items() if value is not None}
