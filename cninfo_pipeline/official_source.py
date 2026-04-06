from __future__ import annotations

import json
import re
import unicodedata
import urllib.request
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

from .client import CninfoClient, CompanyRecord, USER_AGENT
from .paths import ensure_writable_dir


NUMBER_TOKEN_RE = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?")
OFFICIAL_CACHE_SCHEMA_VERSION = 7
STATEMENT_TITLE_KEYWORDS = {
    "balance": "资产负债表",
    "income": "利润表",
    "cash": "现金流量表",
}
STATEMENT_STOP_KEYWORDS = tuple(STATEMENT_TITLE_KEYWORDS.values())
PARENT_STATEMENT_TITLE_KEYWORDS = {
    "balance": "母公司资产负债表",
    "income": "母公司利润表",
    "cash": "母公司现金流量表",
}
LABEL_PREFIXES = ("其中：", "其中:", "加：", "加:", "减：", "减:")
SPECIAL_LABEL_ALIASES: dict[tuple[str, str], dict[str, tuple[str, ...]]] = {
    ("bank", "balance"): {
        "客户存款(吸收存款)": ("客户存款",),
        "其中:同业存放款项": ("同业及其他金融机构存放款项",),
        "买入返售金融资产": ("买入返售款项",),
        "发放贷款及垫款": ("客户贷款及垫款净额",),
    },
}
SPECIAL_LABEL_ALIASES[("bank", "balance")].update(
    {
        "存放同业和其他金融机构款项": ("存放同业及其他金融机构款项", "存放同业款项"),
        "买入返售金融资产": ("买入返售金融资产", "买入返售款项"),
        "贷款和垫款": ("贷款和垫款", "发放贷款及垫款", "客户贷款及垫款净额"),
        "以公允价值计量且其变动计入当期损益的金融投资": ("交易性金融资产",),
        "以摊余成本计量的债务工具投资": ("持有至到期投资", "债权投资"),
        "以公允价值计量且其变动计入其他综合收益的债务工具投资": ("可供出售金融资产", "其他债权投资"),
        "指定为以公允价值计量且其变动计入其他综合收益的权益工具投资": ("其他权益工具投资",),
        "递延所得税资产": ("递延税款借项", "递延所得税资产"),
        "资产合计": ("资产总计",),
        "同业和其他金融机构存放款项": ("同业及其他金融机构存放款项", "其中:同业存放款项"),
        "卖出回购金融资产款": ("卖出回购金融资产款", "卖出回购款项"),
        "客户存款": ("客户存款", "客户存款(吸收存款)"),
        "归属于本行股东权益合计": ("归属于本行股东权益合计", "归属于母公司股东的权益"),
        "股东权益合计": ("股东权益合计", "所有者权益合计"),
        "负债及股东权益总计": ("负债及股东权益总计", "负债及股东权益总额"),
    }
)
SPECIAL_LABEL_ALIASES[("bank", "income")] = {
    "净利息收入": ("净利息收入", "利息净收入"),
    "手续费及佣金收入": ("手续费及佣金收入",),
    "手续费及佣金支出": ("手续费及佣金支出",),
    "净手续费及佣金收入": ("净手续费及佣金收入", "手续费及佣金净收入"),
    "投资收益": ("投资收益", "投资净收益"),
    "其中：对合营企业及联营企业的投资收益": ("其中：对合营企业和联营企业的投资收益", "其中:对联营公司的投资收益"),
    "以摊余成本计量的金融资产终止确认产生的收益": ("以摊余成本计量的金融资产终止确认产生的收益",),
    "公允价值变动收益": ("公允价值变动收益", "公允价值变动净收益"),
    "汇兑净收益": ("汇兑净收益", "汇兑收益"),
    "其他净收入": ("其他净收入",),
    "营业收入小计": ("营业收入小计", "营业收入"),
    "信用减值损失": ("信用减值损失",),
    "其他资产减值损失": ("其他资产减值损失", "资产减值损失"),
    "其他业务成本": ("其他业务成本",),
    "减：所得税费用": ("减：所得税费用", "减:所得税", "减：所得税"),
    "本行股东的净利润": ("本行股东的净利润", "归属于本行股东的净利润", "归属于母公司股东的净利润"),
    "少数股东的净利润": ("少数股东的净利润", "少数股东损益"),
    "基本及稀释每股收益": ("基本及稀释每股收益", "基本每股收益", "稀释每股收益"),
    "权益法下可转损益的其他综合收益": ("权益法下可转损益的其他综合收益",),
    "以公允价值计量且其变动计入其他综合收益的债务工具投资公允价值变动": (
        "以公允价值计量且其变动计入其他综合收益的债务工具投资公允价值变动",
    ),
    "以公允价值计量且其变动计入其他综合收益的债务工具投资信用损失准备": (
        "以公允价值计量且其变动计入其他综合收益的债务工具投资信用损失准备",
    ),
    "外币财务报表折算差额": ("外币财务报表折算差额",),
    "指定为以公允价值计量且其变动计入其他综合收益的权益工具投资公允价值变动": (
        "指定为以公允价值计量且其变动计入其他综合收益的权益工具投资公允价值变动",
    ),
    "重新计量设定受益计划变动额": ("重新计量设定受益计划变动额",),
    "本行股东的综合收益总额": ("本行股东的综合收益总额", "归属于本行股东的综合收益总额", "归属于母公司所有者的综合收益总额"),
    "少数股东的综合收益总额": ("少数股东的综合收益总额", "归属于少数股东的综合收益总额"),
}
SPECIAL_LABEL_ALIASES[("bank", "cash")] = {
    "四、汇率变动对现金及现金等价物的影响额": ("四、汇率变动对现金及现金等价物的影响额", "汇率变动对现金及现金等价物的影响", "汇率变动对现金及现金等价物的影响额"),
    "因：汇率变动及现金及现金等价物的影响额": ("因：汇率变动及现金及现金等价物的影响额", "四、汇率变动对现金及现金等价物的影响额", "汇率变动对现金及现金等价物的影响", "汇率变动对现金及现金等价物的影响额"),
    "加：年初现金及现金等价物余额": ("加：年初现金及现金等价物余额", "加:期初现金及现金等价物余额", "期初现金及现金等价物余额"),
    "六、年末现金及现金等价物余额": ("六、年末现金及现金等价物余额", "六、期末现金及现金等价物余额", "年末现金及现金等价物余额"),
}


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


def extract_statement_values_from_text(
    text: str,
    *,
    template_kind: str,
    statement_type: str,
    requested_labels: Iterable[str],
) -> dict[str, float]:
    occurrence_values = extract_statement_occurrence_values_from_text(
        text,
        template_kind=template_kind,
        statement_type=statement_type,
        requested_labels=requested_labels,
    )
    collapsed: dict[str, float] = {}
    for (label, occurrence), value in occurrence_values.items():
        if occurrence == 1 and label not in collapsed:
            collapsed[label] = value
    return collapsed


def extract_statement_occurrence_values_from_text(
    text: str,
    *,
    template_kind: str,
    statement_type: str,
    requested_labels: Iterable[str],
) -> dict[tuple[str, int], float]:
    lines = _clean_lines(text)
    if not lines:
        return {}

    requested = tuple(label for label in requested_labels if str(label or "").strip())
    multiplier = detect_unit_multiplier(text)
    results: dict[tuple[str, int], float] = {}

    special_values: dict[str, float] = {}
    if template_kind == "company" and statement_type == "income":
        special_values = extract_company_income_values_from_text(text)
    elif template_kind == "bank" and statement_type == "balance":
        special_values = extract_bank_balance_values_from_text(text)

    occurrences: dict[str, int] = {}
    for requested_label in requested:
        label_key = canonicalize_label(requested_label)
        if not label_key:
            continue
        occurrence = occurrences.get(label_key, 0) + 1
        occurrences[label_key] = occurrence
        terms = search_terms_for_label(requested_label, template_kind=template_kind, statement_type=statement_type)
        value = next((special_values[term] for term in terms if term in special_values), None)
        if value is None:
            value = extract_value_for_terms(lines, terms, multiplier=multiplier, occurrence=occurrence)
        if value is not None:
            results[(requested_label, occurrence)] = value
    return results


def extract_value_for_terms(
    lines: list[str],
    terms: Iterable[str],
    *,
    multiplier: int,
    occurrence: int = 1,
) -> float | None:
    for term in terms:
        exact_values = _extract_exact_values(lines, term)
        if len(exact_values) >= occurrence:
            return exact_values[occurrence - 1] * multiplier

    if occurrence != 1:
        return None

    for term in terms:
        value = _extract_first_value(lines, term)
        if value is not None:
            return value * multiplier
    return None


def _extract_exact_values(lines: list[str], term: str) -> list[float]:
    term_key = canonicalize_label(term)
    if not term_key:
        return []

    values: list[float] = []
    for index, line in enumerate(lines):
        if _canonicalize_extracted_label(_line_label_part(line)) != term_key:
            continue
        value = _extract_exact_value(lines, index, term_key)
        if value is not None:
            values.append(value)
    return values


def _extract_exact_value(lines: list[str], index: int, term_key: str) -> float | None:
    line = lines[index]
    label_part = _line_label_part(line)
    if _canonicalize_extracted_label(label_part) != term_key:
        return None

    tail = line[len(label_part) :].strip()
    tokens = NUMBER_TOKEN_RE.findall(tail)
    while tokens and _looks_like_note_number(tokens[0]):
        tokens = tokens[1:]
    for token in tokens:
        value = _parse_number(token)
        if value is not None:
            return value

    for next_index in range(index + 1, min(len(lines), index + 3)):
        next_line = lines[next_index]
        next_label = _line_label_part(next_line)
        next_label_key = _canonicalize_extracted_label(next_label)
        if next_label_key and next_label_key != term_key:
            break
        tokens = NUMBER_TOKEN_RE.findall(next_line)
        while tokens and _looks_like_note_number(tokens[0]):
            tokens = tokens[1:]
        for token in tokens:
            value = _parse_number(token)
            if value is not None:
                return value
    return None


def _line_label_part(line: str) -> str:
    match = NUMBER_TOKEN_RE.search(line)
    if match is None:
        return line.strip()
    return line[: match.start()].strip()


def _canonicalize_extracted_label(value: object | None) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[\(（]\d+[\)）]\s*$", "", text)
    return canonicalize_label(text)


def search_terms_for_label(label: str, *, template_kind: str, statement_type: str) -> tuple[str, ...]:
    raw = str(label or "").strip()
    if not raw:
        return ()

    candidates: list[str] = [raw]
    for prefix in LABEL_PREFIXES:
        if raw.startswith(prefix):
            candidates.append(raw[len(prefix) :].strip())
    if "（" in raw:
        candidates.append(raw.split("（", 1)[0].strip())
    if "(" in raw:
        candidates.append(raw.split("(", 1)[0].strip())

    alias_map = SPECIAL_LABEL_ALIASES.get((template_kind, statement_type), {})
    candidates.extend(alias_map.get(raw, ()))

    unique: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        normalized = canonicalize_label(candidate)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique.append(candidate)
    return tuple(unique)


def official_label_keys(label: str, *, template_kind: str, statement_type: str) -> tuple[str, ...]:
    keys = [canonicalize_label(label)]
    keys.extend(
        canonicalize_label(term)
        for term in search_terms_for_label(label, template_kind=template_kind, statement_type=statement_type)
    )
    return tuple(dict.fromkeys(key for key in keys if key))


def canonicalize_label(value: object | None) -> str:
    text = unicodedata.normalize("NFKC", str(value or "").strip())
    replacements = {
        "\xa0": "",
        " ": "",
        "　": "",
        "（": "",
        "）": "",
        "(": "",
        ")": "",
        "：": "",
        ":": "",
        "、": "",
        ",": "",
        "，": "",
        ";": "",
        "；": "",
        "\n": "",
        "\r": "",
        "\t": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def requested_label_signatures(requested_labels: Iterable[str]) -> tuple[str, ...]:
    occurrences: dict[str, int] = {}
    signatures: list[str] = []
    for label in requested_labels:
        key = canonicalize_label(label)
        if not key:
            continue
        occurrence = occurrences.get(key, 0) + 1
        occurrences[key] = occurrence
        signatures.append(f"{key}#{occurrence}")
    return tuple(signatures)


def serialize_occurrence_values(values: dict[tuple[str, int], float]) -> list[dict[str, object]]:
    payload: list[dict[str, object]] = []
    for (label, occurrence), value in values.items():
        payload.append(
            {
                "label": label,
                "occurrence": occurrence,
                "value": value,
            }
        )
    return payload


def deserialize_occurrence_values(payload: object) -> dict[tuple[str, int], float]:
    if isinstance(payload, dict):
        return {(str(label), 1): value for label, value in payload.items() if isinstance(value, (int, float))}

    if not isinstance(payload, list):
        return {}

    values: dict[tuple[str, int], float] = {}
    for item in payload:
        if not isinstance(item, dict):
            continue
        label = str(item.get("label", "")).strip()
        occurrence = int(item.get("occurrence", 1))
        value = item.get("value")
        if not label or not isinstance(value, (int, float)):
            continue
        values[(label, occurrence)] = value
    return values


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
        requested_labels: Iterable[str] | None = None,
    ) -> dict[tuple[str, int], float]:
        year = period_end[:4]
        cache_path = self.cache_dir / f"{company.seccode}_{year}_{template_kind}_{statement_type}.json"
        requested = tuple(label for label in requested_labels or () if str(label or "").strip())
        requested_signatures = requested_label_signatures(requested)

        if cache_path.exists():
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            if (
                isinstance(payload, dict)
                and payload.get("schema_version") == OFFICIAL_CACHE_SCHEMA_VERSION
                and set(requested_signatures).issubset(set(payload.get("requested_signatures", [])))
            ):
                return deserialize_occurrence_values(payload.get("values"))

        pdf_url = self._find_annual_report_pdf_url(company.seccode, year)
        if not pdf_url:
            cache_path.write_text(
                json.dumps(
                    {
                        "schema_version": OFFICIAL_CACHE_SCHEMA_VERSION,
                        "requested_signatures": list(requested_signatures),
                        "values": [],
                    },
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
            requested_labels=requested,
        )
        cache_path.write_text(
            json.dumps(
                {
                    "schema_version": OFFICIAL_CACHE_SCHEMA_VERSION,
                    "requested_signatures": list(requested_signatures),
                    "values": serialize_occurrence_values(values),
                },
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

    def _extract_statement_values(
        self,
        pdf_path: Path,
        *,
        template_kind: str,
        statement_type: str,
        requested_labels: Iterable[str] | None = None,
    ) -> dict[tuple[str, int], float]:
        if template_kind == "company" and statement_type == "income":
            statement_text = self._extract_company_income_text(pdf_path)
        else:
            statement_text = self._extract_statement_text(pdf_path, statement_type=statement_type)

        if not statement_text:
            return {}

        results = extract_statement_occurrence_values_from_text(
            statement_text,
            template_kind=template_kind,
            statement_type=statement_type,
            requested_labels=requested_labels or (),
        )
        if template_kind == "bank" and statement_type == "balance":
            special_values = self._extract_bank_balance_values(pdf_path)
            occurrences: dict[str, int] = {}
            for requested_label in requested_labels or ():
                label_key = canonicalize_label(requested_label)
                if not label_key:
                    continue
                occurrence = occurrences.get(label_key, 0) + 1
                occurrences[label_key] = occurrence
                result_key = (requested_label, occurrence)
                if result_key in results:
                    continue
                for term in search_terms_for_label(
                    requested_label,
                    template_kind=template_kind,
                    statement_type=statement_type,
                ):
                    if term in special_values:
                        results[result_key] = special_values[term]
                        break
        return results

    def _extract_company_income_text(self, pdf_path: Path) -> str:
        reader = PdfReader(str(pdf_path))
        best_score = float("-inf")
        best_text = ""
        scan_limit = min(len(reader.pages), 220)

        for page_index in range(scan_limit):
            page_text = reader.pages[page_index].extract_text() or ""
            if "利润表" not in page_text and "合并利润表" not in page_text:
                continue

            window = [page_text]
            for next_index in range(page_index + 1, min(scan_limit, page_index + 3)):
                window.append(reader.pages[next_index].extract_text() or "")
            combined = "\n".join(window)
            for stop_marker in ("母公司利润表", "母公司现金流量表", "合并现金流量表"):
                if stop_marker in combined:
                    combined = combined.split(stop_marker, 1)[0]

            score = 0
            if "合并利润表" in page_text:
                score += 100
            elif "合并利润表" in combined:
                score += 80

            if "第十节 财务报告" in combined or "财务报告" in page_text:
                score += 20

            for marker in ("营业收入", "营业总成本", "营业利润", "利润总额", "净利润"):
                if marker in combined:
                    score += 8

            for bad_marker in ("主营业务分析", "变动分析表", "相关科目变动分析表", "管理层讨论与分析"):
                if bad_marker in combined:
                    score -= 60

            if score > best_score:
                best_score = score
                best_text = combined

        if best_score >= 40:
            return best_text
        return self._extract_statement_text(pdf_path, statement_type="income")

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

    def _extract_statement_text(self, pdf_path: Path, *, statement_type: str) -> str:
        reader = PdfReader(str(pdf_path))
        keyword = STATEMENT_TITLE_KEYWORDS[statement_type]
        parent_keyword = PARENT_STATEMENT_TITLE_KEYWORDS[statement_type]
        candidates: list[tuple[int, int]] = []
        for index, page in enumerate(reader.pages[:220]):
            page_text = page.extract_text() or ""
            if keyword not in page_text:
                continue
            score = 2 if "合并" in page_text else 0
            candidates.append((score, index))

        if not candidates:
            return ""

        candidates.sort(key=lambda item: (-item[0], item[1]))
        start_index = candidates[0][1]
        window: list[str] = []
        for index in range(start_index, min(len(reader.pages), start_index + 5)):
            page_text = reader.pages[index].extract_text() or ""
            if index > start_index and (
                parent_keyword in page_text
                or any(title in page_text for title in STATEMENT_STOP_KEYWORDS if title != keyword)
            ):
                break
            window.append(page_text)
        return "\n".join(window)


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
    if re.fullmatch(r"\d{1,2}", stripped):
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
