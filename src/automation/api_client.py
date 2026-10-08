"""Small REST client: auth header, timeouts, automatic retries, optional pagination."""
from __future__ import annotations

import logging
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from .config import ApiSettings, ConfigError, get_api_settings
from .io_files import dig

log = logging.getLogger(__name__)


class ApiError(Exception):
    pass


class ApiClient:
    def __init__(self, settings: ApiSettings | None = None, session: requests.Session | None = None):
        self.settings = settings or get_api_settings()
        self.session = session or requests.Session()
        retry = Retry(total=self.settings.max_retries, backoff_factor=1,
                      status_forcelist=(429, 500, 502, 503, 504), respect_retry_after_header=True)
        for prefix in ("http://", "https://"):
            self.session.mount(prefix, HTTPAdapter(max_retries=retry))
        self.session.headers["Accept"] = "application/json"
        if self.settings.api_key:
            value = f"{self.settings.key_prefix} {self.settings.api_key}".strip()
            self.session.headers[self.settings.key_header] = value

    def url(self, path_or_url: str) -> str:
        if path_or_url.lower().startswith(("http://", "https://")):
            return path_or_url
        if not self.settings.base_url:
            raise ConfigError("Give a full URL or set API_BASE_URL in .env")
        return f"{self.settings.base_url.rstrip('/')}/{path_or_url.lstrip('/')}"

    def request(self, method: str, path_or_url: str, params: dict[str, Any] | None = None,
                json: Any = None) -> requests.Response:
        url = self.url(path_or_url)
        log.info("%s %s", method.upper(), url)  # API key is never logged
        try:
            resp = self.session.request(method.upper(), url, params=params, json=json,
                                        timeout=self.settings.timeout, verify=self.settings.verify)
        except requests.RequestException as e:
            raise ApiError(f"{method.upper()} {url} failed: {e}") from e
        if not resp.ok:
            raise ApiError(f"{method.upper()} {url} -> HTTP {resp.status_code}: {resp.text[:300]}")
        return resp

    def get_json(self, path_or_url: str, params: dict[str, Any] | None = None) -> Any:
        return self.request("GET", path_or_url, params=params).json()

    def post_json(self, path_or_url: str, payload: Any, params: dict[str, Any] | None = None) -> Any:
        return self.request("POST", path_or_url, params=params, json=payload).json()

    def get_all_pages(self, path_or_url: str, params: dict[str, Any] | None = None, page_param: str = "page",
                      records_path: str | None = None, start_page: int = 1, max_pages: int = 50) -> list[Any]:
        """Call ?page=1,2,3... until an empty page comes back (or max_pages is reached)."""
        records: list[Any] = []
        for page in range(start_page, start_page + max_pages):
            data = dig(self.get_json(path_or_url, {**(params or {}), page_param: page}), records_path)
            if not data:
                break
            records.extend(data if isinstance(data, list) else [data])
        return records
