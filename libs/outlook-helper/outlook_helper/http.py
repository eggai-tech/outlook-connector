"""The single HTTP seam to Microsoft Graph.

``GraphSession`` is the only place that talks to Graph. It injects the bearer
token, asks Graph for immutable ids, maps non-2xx responses to
:class:`GraphError`, retries throttling/5xx responses honouring ``Retry-After``,
lazily follows pagination links, and streams large downloads to disk.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Callable, Iterator

import httpx

from outlook_helper.auth import Credential
from outlook_helper.exceptions import GraphError

DEFAULT_BASE_URL = "https://graph.microsoft.com/v1.0"
_RETRY_STATUS = frozenset({429, 503})
_REQUEST_ID_HEADERS = ("request-id", "client-request-id", "x-ms-request-id")

#: Asks Graph to return immutable ids. Without it, the ``id`` of an Outlook item
#: encodes its current folder, so it changes whenever the item moves (a user
#: filing a mail, a rule firing) and is useless as a durable key. With it, ``id``
#: survives moves. Ids handed back this way work in later request URLs whether or
#: not the header is set on those requests, so we send it on every request.
IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"'


class GraphSession:
    def __init__(
        self,
        credential: Credential,
        *,
        base_url: str = DEFAULT_BASE_URL,
        max_retries: int = 3,
        client: httpx.Client | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._credential = credential
        self._base_url = base_url.rstrip("/")
        self._max_retries = max_retries
        self._client = client or httpx.Client()
        self._sleep = sleep

    # --- low-level ---

    def _url(self, path: str) -> str:
        if path.startswith("http://") or path.startswith("https://"):
            return path
        return f"{self._base_url}/{path.lstrip('/')}"

    def _auth_headers(self, extra: dict | None) -> dict:
        headers = {"Authorization": f"Bearer {self._credential.get_token()}"}
        if extra:
            headers.update(extra)
        headers["Prefer"] = _prefer_immutable_ids(_pop_prefer(headers))
        return headers

    def request(
        self,
        method: str,
        path: str,
        *,
        params: dict | None = None,
        json: Any | None = None,
        content: bytes | None = None,
        headers: dict | None = None,
    ) -> httpx.Response:
        url = self._url(path)
        attempt = 0
        while True:
            response = self._client.request(
                method,
                url,
                params=params,
                json=json,
                content=content,
                headers=self._auth_headers(headers),
            )
            if response.status_code in _RETRY_STATUS and attempt < self._max_retries:
                self._sleep(_retry_after_seconds(response, attempt))
                attempt += 1
                continue
            if response.status_code >= 400:
                raise _to_graph_error(response)
            return response

    def get_json(
        self, path: str, params: dict | None = None, *, headers: dict | None = None
    ) -> dict:
        return self.request("GET", path, params=params, headers=headers).json()

    def get_text(
        self, path: str, params: dict | None = None, *, headers: dict | None = None
    ) -> str:
        """Return the response body for ``path`` as text (e.g. ``$value`` MIME)."""
        return self.request("GET", path, params=params, headers=headers).text

    def paginate(
        self, path: str, params: dict | None = None, *, headers: dict | None = None
    ) -> Iterator[dict]:
        """Yield items across pages, following ``@odata.nextLink`` on demand."""
        page = self.get_json(path, params, headers=headers)
        while True:
            yield from page.get("value", [])
            next_link = page.get("@odata.nextLink")
            if not next_link:
                return
            page = self.get_json(next_link, headers=headers)  # nextLink carries its own query

    def download(self, path: str, dest: Path | str) -> Path:
        """Stream the response body for ``path`` to ``dest`` and return the path."""
        dest = Path(dest)
        url = self._url(path)
        with self._client.stream(
            "GET", url, headers=self._auth_headers(None)
        ) as response:
            if response.status_code >= 400:
                response.read()
                raise _to_graph_error(response)
            with dest.open("wb") as fh:
                for chunk in response.iter_bytes():
                    fh.write(chunk)
        return dest

    def upload_chunk(
        self, upload_url: str, data: bytes, content_range: str
    ) -> httpx.Response:
        """PUT one chunk to an upload session URL.

        The upload URL is pre-authenticated, so no bearer token is attached.
        """
        response = self._client.put(
            upload_url,
            content=data,
            headers={"Content-Range": content_range},
        )
        if response.status_code >= 400:
            raise _to_graph_error(response)
        return response

    def close(self) -> None:
        self._client.close()


def _pop_prefer(headers: dict) -> str | None:
    """Remove and return any caller-supplied ``Prefer`` value (whatever casing)."""
    for key in [k for k in headers if k.lower() == "prefer"]:
        return headers.pop(key)
    return None


def _prefer_immutable_ids(existing: str | None) -> str:
    """Prepend the immutable-id preference to ``existing``, a Prefer value list."""
    if not existing:
        return IMMUTABLE_ID_PREFERENCE
    others = [
        p.strip()
        for p in existing.split(",")
        if p.strip() and p.strip() != IMMUTABLE_ID_PREFERENCE
    ]
    return ", ".join([IMMUTABLE_ID_PREFERENCE, *others])


def _retry_after_seconds(response: httpx.Response, attempt: int) -> float:
    header = response.headers.get("Retry-After")
    if header is not None:
        try:
            return float(header)
        except ValueError:
            pass
    return float(2**attempt)  # exponential backoff fallback


def _to_graph_error(response: httpx.Response) -> GraphError:
    code: str | None = None
    message = response.text
    try:
        payload = response.json()
    except ValueError:
        payload = None
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            code = error.get("code")
            message = error.get("message", message)
        elif isinstance(error, str):
            code = error
            message = payload.get("error_description", message)
    request_id = None
    for header in _REQUEST_ID_HEADERS:
        if header in response.headers:
            request_id = response.headers[header]
            break
    return GraphError(
        status_code=response.status_code,
        message=message,
        code=code,
        request_id=request_id,
    )
