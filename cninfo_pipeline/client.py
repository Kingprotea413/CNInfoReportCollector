from __future__ import annotations

import ctypes
import json
import os
import re
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import quickjs

from .paths import ensure_writable_dir, resolve_default_cache_dir


BASE_URL = "http://webapi.cninfo.com.cn"
REFERER = "http://webapi.cninfo.com.cn/shgs/company4.html"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
FILE_ATTRIBUTE_HIDDEN = 0x02


class CninfoError(RuntimeError):
    """Raised when CNInfo responds with an error."""


@dataclass(frozen=True)
class CompanyRecord:
    seccode: str
    secname: str
    orgname: str

    @classmethod
    def from_api(cls, payload: dict[str, Any]) -> "CompanyRecord":
        return cls(
            seccode=str(payload["SECCODE"]),
            secname=str(payload["SECNAME"]),
            orgname=str(payload["ORGNAME"]),
        )


class EncKeyProvider:
    """Executes CNInfo's own CryptoJS bundle to build Accept-EncKey."""

    def __init__(self, opener: urllib.request.OpenerDirector, cache_dir: Path) -> None:
        self._opener = opener
        self._cache_path = cache_dir / "crypto-js.js"
        self._context: quickjs.Context | None = None

    def get(self) -> str:
        if self._context is None:
            self._context = quickjs.Context()
            self._context.eval(self._bootstrap_script())
            self._context.eval(self._load_bundle())
        return str(self._context.eval("indexcode.getResCode()"))

    def _load_bundle(self) -> str:
        if self._cache_path.exists():
            return self._cache_path.read_text(encoding="utf-8")

        request = urllib.request.Request(
            f"{BASE_URL}/js/crypto-js.js",
            headers={"User-Agent": USER_AGENT, "Referer": REFERER},
        )
        script = self._opener.open(request, timeout=30).read().decode("utf-8")
        ensure_cache_dir(self._cache_path.parent)
        self._cache_path.write_text(script, encoding="utf-8")
        return script

    @staticmethod
    def _bootstrap_script() -> str:
        return """
var window = globalThis;
var noop = function() {};
var body = { appendChild: function(node) { return node; }, removeChild: noop };
body.parentNode = body;
var document = {
  body: body,
  head: body,
  documentElement: body,
  currentScript: { parentNode: body },
  createElement: function() { return { setAttribute: noop, style: {}, parentNode: body }; },
  getElementsByTagName: function() { return [body]; },
  querySelector: function() { return body; }
};
var localStorage = {
  _store: {},
  getItem: function(key) {
    return Object.prototype.hasOwnProperty.call(this._store, key) ? this._store[key] : null;
  },
  setItem: function(key, value) {
    this._store[key] = String(value);
  }
};
function btoa(input) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
  var str = String(input);
  var output = '';
  for (var block = 0, charCode, idx = 0, map = chars;
       str.charAt(idx | 0) || (map = '=', idx % 1);
       output += map.charAt(63 & block >> 8 - idx % 1 * 8)) {
    charCode = str.charCodeAt(idx += 3 / 4);
    if (charCode > 0xFF) {
      throw new Error('Invalid character');
    }
    block = block << 8 | charCode;
  }
  return output;
}
"""


class CninfoClient:
    def __init__(self, cache_dir: str | Path | None = None) -> None:
        self.opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        self.set_cache_dir(cache_dir or resolve_default_cache_dir())

    def set_cache_dir(self, cache_dir: str | Path) -> None:
        self.cache_dir = ensure_writable_dir(cache_dir)
        ensure_cache_dir(self.cache_dir)
        self.enc_key_provider = EncKeyProvider(self.opener, self.cache_dir)
        self._company_cache_path = self.cache_dir / "companies.json"

    def fetch_company_catalog(self, use_cache: bool = True) -> list[CompanyRecord]:
        if use_cache and self._company_cache_path.exists():
            payload = json.loads(self._company_cache_path.read_text(encoding="utf-8"))
        else:
            payload = self._request_json("/api/sysapi/p_sysapi1067")
            ensure_cache_dir(self._company_cache_path.parent)
            self._company_cache_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

        records = payload.get("records", [])
        return [CompanyRecord.from_api(item) for item in records]

    def search_company(self, query: str) -> CompanyRecord:
        query = query.strip()
        if not query:
            raise ValueError("公司名称或证券代码不能为空。")

        companies = self.fetch_company_catalog()
        exact_code = [item for item in companies if item.seccode == query]
        if exact_code:
            return exact_code[0]

        normalized = query.casefold()
        exact_name = [
            item
            for item in companies
            if item.secname.casefold() == normalized or item.orgname.casefold() == normalized
        ]
        if exact_name:
            return exact_name[0]

        fuzzy = [
            item
            for item in companies
            if normalized in item.secname.casefold()
            or normalized in item.orgname.casefold()
            or normalized in item.seccode.casefold()
        ]
        if len(fuzzy) == 1:
            return fuzzy[0]
        if not fuzzy:
            raise ValueError(f"没有找到与“{query}”匹配的公司。")

        options = ", ".join(f"{item.secname}({item.seccode})" for item in fuzzy[:8])
        raise ValueError(f"匹配到多家公司，请输入更精确的名称或代码：{options}")

    def fetch_balance_sheet(self, seccode: str) -> list[dict[str, Any]]:
        payload = self._request_json("/api/stock/p_stock2300", {"scode": seccode})
        return list(payload.get("records", []))

    def fetch_income_statement(self, seccode: str) -> list[dict[str, Any]]:
        payload = self._request_json("/api/stock/p_stock2301", {"scode": seccode})
        return list(payload.get("records", []))

    def fetch_cash_flow_statement(self, seccode: str) -> list[dict[str, Any]]:
        payload = self._request_json("/api/stock/p_stock2302", {"scode": seccode})
        return list(payload.get("records", []))

    def _request_json(
        self,
        path: str,
        params: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = f"{BASE_URL}{path}"
        if params:
            query = urllib.parse.urlencode(params)
            url = f"{url}?{query}"

        request = urllib.request.Request(url, headers=self._headers())
        with self.opener.open(request, timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8"))

        result_code = payload.get("resultcode")
        if result_code not in (None, 200):
            raise CninfoError(payload.get("resultmsg", f"CNInfo request failed: {result_code}"))
        return payload

    def _headers(self) -> dict[str, str]:
        return {
            "User-Agent": USER_AGENT,
            "Referer": REFERER,
            "Accept-EncKey": self.enc_key_provider.get(),
        }


def natural_sort_key(value: str) -> list[Any]:
    parts = re.split(r"(\d+)", value)
    key: list[Any] = []
    for part in parts:
        if part.isdigit():
            key.append(int(part))
        else:
            key.append(part)
    return key


def ensure_cache_dir(cache_dir: Path) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    if os.name != "nt":
        return

    try:
        existing = ctypes.windll.kernel32.GetFileAttributesW(str(cache_dir))
        if existing != -1 and not existing & FILE_ATTRIBUTE_HIDDEN:
            ctypes.windll.kernel32.SetFileAttributesW(str(cache_dir), existing | FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass
