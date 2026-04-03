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
from deli_api.schemas import CaseHit, LawHit


SUCCESS_CODES = {0, "0", 200, "200", "success", "SUCCESS"}
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
    """Thin client for fixed Deli API response template:
    {success, code, msg, body: {data, queryId, totalPage, totalCount}}
    """

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
        load_dotenv(
            dotenv_path=os.fspath(dotenv_path) if dotenv_path is not None else None,
            override=override,
        )
        app_id = os.getenv("DELILEGAL_APP_ID")
        secret = os.getenv("DELILEGAL_SECRET")
        base_url = os.getenv("DELILEGAL_BASE_URL", "https://openapi.delilegal.com")
        timeout = float(os.getenv("DELILEGAL_TIMEOUT", "30"))

        missing = [
            key
            for key, value in (
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
    ) -> list[LawHit]:
        condition: dict[str, Any] = {
            "keywords": [query],
            "fieldName": field_name,
        }
        if time_liness_type_arr:
            condition["timeLinessTypeArr"] = list(time_liness_type_arr)
        if publish_year_start is not None:
            condition["publishYearStart"] = publish_year_start
        if publish_year_end is not None:
            condition["publishYearEnd"] = publish_year_end
        if active_year_start is not None:
            condition["activeYearStart"] = active_year_start
        if active_year_end is not None:
            condition["activeYearEnd"] = active_year_end
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
    ) -> list[CaseHit]:
        condition: dict[str, Any] = {"keywordArr": [query]}
        if case_year_start is not None:
            condition["caseYearStart"] = case_year_start
        if case_year_end is not None:
            condition["caseYearEnd"] = case_year_end
        if court_level_arr:
            condition["courtLevelArr"] = list(court_level_arr)
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

        if not isinstance(data, dict):
            raise DeliLegalResponseFormatError(
                "Unexpected response format: expected top-level object.",
                path=path,
                payload=payload,
                response_text=_truncate_text(json.dumps(data, ensure_ascii=False)),
            )
        body = data.get("body")
        if not isinstance(body, dict):
            raise DeliLegalResponseFormatError(
                "Unexpected response format: expected body object.",
                path=path,
                payload=payload,
                response_text=_truncate_text(json.dumps(data, ensure_ascii=False)),
            )
        items = body.get("data")
        if not isinstance(items, list):
            raise DeliLegalResponseFormatError(
                "Unexpected response format: expected body.data list.",
                path=path,
                payload=payload,
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
            response_payload.get("msg")
            or response_payload.get("message")
            or "Unknown upstream error."
        )

        # Known success template: success=True and code in {0, 200, ...}
        if success is True and code in SUCCESS_CODES:
            return
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
                details={
                    "success": success,
                    "query_id": _extract_query_id(response_payload),
                },
                response_text=_truncate_text(json.dumps(response_payload, ensure_ascii=False)),
            )

    def _normalize_law_hit(self, item: dict[str, Any]) -> LawHit:
        title = _to_str(item.get("title")) or "Untitled law"
        issued_no = _to_str(item.get("issuedNo"))
        content = _to_str(item.get("content")) or title
        citation = f"{title}{issued_no}" if issued_no else title

        return LawHit(
            title=title,
            content=content,
            score=_to_float(item.get("score")),
            source_id=_to_str(item.get("id")),
            citation=citation,
            url=_to_str(item.get("url")),
            law_name=title,
            article_no=issued_no,
            publisher_name=_to_str(item.get("publisherName")),
            level_name=_to_str(item.get("levelName")),
            publish_date=_to_str(item.get("publishDate")),
            active_date=_to_str(item.get("activeDate")),
            timeliness_type=_to_str(item.get("timelinessName")),
            raw=item,
        )

    def _normalize_case_hit(self, item: dict[str, Any]) -> CaseHit:
        case_no = _to_str(item.get("caseNumber"))
        title = _to_str(item.get("title")) or case_no or "Untitled case"
        content = _to_str(item.get("content")) or title
        court_name = _to_str(item.get("court"))

        return CaseHit(
            title=title,
            content=content,
            score=_to_float(item.get("score")),
            source_id=_to_str(item.get("id")),
            citation=case_no or title,
            url=_to_str(item.get("url")),
            case_no=case_no,
            court_name=court_name,
            case_type=_to_str(item.get("caseType")),
            cause=_to_str(item.get("cause")),
            case_date=_to_str(item.get("judgementDate")),
            level_of_trial=_to_str(item.get("levelOfTrial")),
            judgement_type=_to_str(item.get("judgementType")),
            publish_type=_to_str(item.get("publishType")),
            publish_type_name=_to_str(item.get("publishTypeName")),
            raw=item,
        )


def init_case_client(
    dotenv_path: str | os.PathLike[str] | None = None,
    *,
    override: bool = False,
) -> DeliLegalClient:
    """One-line factory for creating a DeliLegal client from env vars."""
    return DeliLegalClient.from_env(dotenv_path=dotenv_path, override=override)


def _extract_query_id(payload: Mapping[str, Any]) -> str | None:
    body = payload.get("body")
    if isinstance(body, Mapping):
        query_id = body.get("queryId") or body.get("query_id")
        return _to_str(query_id)
    return _to_str(payload.get("queryId") or payload.get("query_id"))


def _is_auth_error(*, code: Any, message: str) -> bool:
    if code in AUTH_ERROR_CODES:
        return True
    normalized = message.lower()
    return any(token in normalized for token in AUTH_ERROR_TOKENS)


def _to_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _truncate_text(text: str | None, limit: int = 2000) -> str | None:
    if text is None:
        return None
    if len(text) <= limit:
        return text
    return f"{text[:limit]}...<truncated>"
