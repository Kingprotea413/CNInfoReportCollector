from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

from .client import CninfoClient, CompanyRecord, USER_AGENT
from .paths import ensure_writable_dir


NUMBER_TOKEN_RE = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?")
OFFICIAL_CACHE_SCHEMA_VERSION = 1


def detect_unit_multiplier(text: str) -> int:
    normalized = re.sub(r"\s+", "", text)
    if "百万元" in normalized:
        return 1_000_000
    if "万元" in normalized:
        return 10_000
    if "千元" in normalized:
        return 1_000
    return 1


def extract_company_income_values_from_text(text: str) -> dict[str, float]:
    lines = _clean_lines(text)
    if not lines:
        return {}

    finance_index = next((idx for idx, line in enumerate(lines) if "财务费用" in line), 0)
    search_lines = lines[finance_index : finance_index + 24] if finance_index < len(lines) else lines
    multiplier = detect_unit_multiplier(text)
    results: dict[str, float] = {}

    labels = ("利息费用", "利息收入", "其他收益", "信用减值损失", "资产减值损失", "资产处置收益")
    for label in labels:
        value = _extract_first_value(search_lines, label, stop_labels=labels)
        if value is not None:
            results[label] = value * multiplier
    return results


def extract_bank_balance_values_from_text(text: str, *, multiplier: int | None = None) -> dict[str, float]:
    lines = _clean_lines(text)
    if not lines:
        return {}

    scale = multiplier or detect_unit_multiplier(text)
    results: dict[str, float] = {}
    label_map = {
        "客户存款(吸收存款)": "客户存款",
        "其中:同业存放款项": "同业及其他金融机构存放款项",
        "拆入资金": "拆入资金",
        "发放贷款及垫款": "客户贷款及垫款净额",
        "现金及存放中央银行款项": "现金及存放中央银行款项",
        "买入返售金融资产": "买入返售款项",
    }

    for export_label, search_term in label_map.items():
        value = _extract_first_value(lines, search_term)
        if value is not None:
            results[export_label] = value * scale
    return results


class OfficialAnnualReportSource:
    def __init__(self, client: CninfoClient) -> None:
        self.client = client
        self.cache_dir = ensure_writable_dir(Path(self.client.cache_dir) / "official_reports")

    def get_statement_overrides(
        self,
        company: CompanyRecord,
        *,
        template_kind: str,
        statement_type: str,
        period_end: str,
    ) -> dict[str, float]:
        year = period_end[:4]
        cache_path = self.cache_dir / f"{company.seccode}_{year}_{template_kind}_{statement_type}.json"
        if cache_path.exists():
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            if (
                isinstance(payload, dict)
                and payload.get("schema_version") == OFFICIAL_CACHE_SCHEMA_VERSION
                and isinstance(payload.get("values"), dict)
                and payload.get("values")
            ):
                return payload["values"]

        pdf_url = self._find_annual_report_pdf_url(company.seccode, year)
        if not pdf_url:
            cache_path.write_text(
                json.dumps(
                    {"schema_version": OFFICIAL_CACHE_SCHEMA_VERSION, "values": {}},
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            return {}

        pdf_path = self._download_pdf(company.seccode, year, pdf_url)
        values = self._extract_statement_values(
            pdf_path,
            template_kind=template_kind,
            statement_type=statement_type,
        )
        cache_path.write_text(
            json.dumps(
                {"schema_version": OFFICIAL_CACHE_SCHEMA_VERSION, "values": values},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return values

    def _announcement_records(self, seccode: str) -> list[dict]:
        cache_path = self.cache_dir / f"announcements_{seccode}.json"
        if cache_path.exists():
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
        else:
            payload = self.client._request_json("/api/sysapi/p_sysapi1091", {"stype": 1, "scode": seccode})
            cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return list(payload.get("records", []))

    def _find_annual_report_pdf_url(self, seccode: str, year: str) -> str | None:
        records = self._announcement_records(seccode)
        candidates: list[tuple[str, str]] = []
        for record in records:
            title = str(record.get("F002V", ""))
            sort_name = str(record.get("SORTNAME", ""))
            url = str(record.get("F003V", "")).strip()
            if sort_name != "定期报告" or not url.lower().endswith(".pdf"):
                continue
            if year not in title:
                continue
            if "摘要" in title or "英文" in title:
                continue
            if "年度报告" not in title and "年报" not in title:
                continue
            declare_date = str(record.get("F001D", ""))
            candidates.append((declare_date, url))

        if not candidates:
            return None
        candidates.sort(reverse=True)
        return candidates[0][1]

    def _download_pdf(self, seccode: str, year: str, url: str) -> Path:
        pdf_path = self.cache_dir / f"{seccode}_{year}.pdf"
        if pdf_path.exists():
            return pdf_path

        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Referer": url})
        with self.client.opener.open(request, timeout=120) as response:
            pdf_path.write_bytes(response.read())
        return pdf_path

    def _extract_statement_values(self, pdf_path: Path, *, template_kind: str, statement_type: str) -> dict[str, float]:
        if template_kind == "company" and statement_type == "income":
            return self._extract_company_income_values(pdf_path)
        if template_kind == "bank" and statement_type == "balance":
            return self._extract_bank_balance_values(pdf_path)
        return {}

    def _extract_company_income_values(self, pdf_path: Path) -> dict[str, float]:
        reader = PdfReader(str(pdf_path))
        for page_index, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            if "利润表" not in page_text:
                continue
            window = [page_text]
            if page_index + 1 < len(reader.pages):
                window.append(reader.pages[page_index + 1].extract_text() or "")
            combined = "\n".join(window)
            if "财务费用" not in combined:
                continue
            values = extract_company_income_values_from_text(combined)
            if values:
                return values
        return {}

    def _extract_bank_balance_values(self, pdf_path: Path) -> dict[str, float]:
        reader = PdfReader(str(pdf_path))
        page_texts = [(page.extract_text() or "") for page in reader.pages[:60]]
        document_multiplier = next((m for m in (detect_unit_multiplier(text) for text in page_texts) if m != 1), 1)
        results: dict[str, float] = {}

        for text in page_texts:
            if not results and "客户存款" in text and "同业及其他金融机构存放款项" in text and "拆入资金" in text:
                results.update(extract_bank_balance_values_from_text(text, multiplier=document_multiplier))
                continue

            if "客户贷款及垫款净额" in text and "现金及存放中央银行款项" in text:
                results.update(extract_bank_balance_values_from_text(text, multiplier=document_multiplier))

            if {
                "客户存款(吸收存款)",
                "其中:同业存放款项",
                "拆入资金",
                "发放贷款及垫款",
                "现金及存放中央银行款项",
                "买入返售金融资产",
            }.issubset(results):
                break

        return results


def _clean_lines(text: str) -> list[str]:
    return [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]


def _extract_first_value(lines: Iterable[str], label: str, *, stop_labels: Iterable[str] = ()) -> float | None:
    lines = list(lines)
    for index in range(len(lines)):
        window = " ".join(lines[index : index + 3])
        if label not in window:
            continue

        tail = window.split(label, 1)[1]
        stop_positions = [
            tail.find(stop_label)
            for stop_label in stop_labels
            if stop_label != label and stop_label in tail
        ]
        if stop_positions:
            tail = tail[: min(stop_positions)]
        tokens = NUMBER_TOKEN_RE.findall(tail)
        if tokens and _looks_like_note_number(tokens[0]):
            tokens = tokens[1:]
        for token in tokens:
            value = _parse_number(token)
            if value is not None:
                return value
    return None


def _looks_like_note_number(token: str) -> bool:
    stripped = token.strip("()")
    if re.fullmatch(r"\d{1,3}", stripped):
        return True
    if re.fullmatch(r"20\d{2}", stripped):
        return True
    return False


def _parse_number(token: str) -> float | None:
    negative = token.startswith("(") and token.endswith(")")
    stripped = token.strip("()").replace(",", "")
    if stripped in {"", "-", "—"}:
        return None
    value = float(stripped)
    return -value if negative else value
