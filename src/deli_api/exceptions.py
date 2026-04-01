from __future__ import annotations

from typing import Any


class DeliLegalError(Exception):
    """Base exception for DeliLegal integrations."""

    def __init__(
        self,
        message: str,
        *,
        code: str | int | None = None,
        status_code: int | None = None,
        path: str | None = None,
        details: dict[str, Any] | None = None,
        payload: dict[str, Any] | None = None,
        response_text: str | None = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.code = code
        self.status_code = status_code
        self.path = path
        self.details = details or {}
        self.payload = payload
        self.response_text = response_text

    def to_dict(self) -> dict[str, Any]:
        return {
            "message": self.message,
            "code": self.code,
            "status_code": self.status_code,
            "path": self.path,
            "details": self.details,
            "payload": self.payload,
            "response_text": self.response_text,
        }

    def __str__(self) -> str:
        parts = [self.message]
        if self.code is not None:
            parts.append(f"code={self.code}")
        if self.status_code is not None:
            parts.append(f"status_code={self.status_code}")
        if self.path:
            parts.append(f"path={self.path}")
        return " | ".join(parts)


class DeliLegalConfigError(DeliLegalError):
    """Raised when local configuration is incomplete or invalid."""


class DeliLegalRequestError(DeliLegalError):
    """Raised when the HTTP request cannot be completed."""


class DeliLegalHTTPStatusError(DeliLegalRequestError):
    """Raised when the upstream service returns a non-2xx HTTP status."""


class DeliLegalResponseDecodeError(DeliLegalError):
    """Raised when the upstream response is not valid JSON."""


class DeliLegalResponseFormatError(DeliLegalError):
    """Raised when the upstream JSON schema is missing required fields."""


class DeliLegalAPIError(DeliLegalError):
    """Raised when the upstream API reports a business-level error."""


class DeliLegalAuthenticationError(DeliLegalAPIError):
    """Raised when appid/secret authentication fails."""


class DeliLegalUpstreamError(DeliLegalAPIError):
    """Raised for non-auth upstream business errors."""


__all__ = [
    "DeliLegalAPIError",
    "DeliLegalAuthenticationError",
    "DeliLegalConfigError",
    "DeliLegalError",
    "DeliLegalHTTPStatusError",
    "DeliLegalRequestError",
    "DeliLegalResponseDecodeError",
    "DeliLegalResponseFormatError",
    "DeliLegalUpstreamError",
]
