from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from openpyxl import load_workbook

from .paths import resolve_project_root


TEMPLATE_GLOB_PATTERNS = (
    "cninfo_pipeline/*模版*.xlsx",
    "cninfo_pipeline/*模板*.xlsx",
    "cninfo_pipeline/templates/*.xlsx",
)


@dataclass(frozen=True)
class TemplateSpec:
    template_id: str
    display_name: str
    kind: str
    path: Path


def template_data_files(project_root: Path | None = None) -> list[tuple[str, str]]:
    root = project_root or resolve_project_root()
    datas: list[tuple[str, str]] = []
    seen: set[Path] = set()

    for pattern in TEMPLATE_GLOB_PATTERNS:
        for path in sorted(root.glob(pattern)):
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            destination = path.relative_to(root).parent.as_posix()
            datas.append((str(path), destination))

    return datas


def _normalize_marker(value: str) -> str:
    return "".join(str(value).replace("\xa0", "").split())


def _guess_template_kind(path: Path) -> str:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        markers: set[str] = set()
        for sheet in workbook.worksheets[:3]:
            for row in range(1, min(sheet.max_row, 25) + 1):
                markers.add(_normalize_marker(str(sheet.cell(row, 1).value or "")))
        joined_titles = " ".join(sheet.title for sheet in workbook.worksheets)
    finally:
        workbook.close()

    bank_markers = {
        "现金及存放中央银行款项",
        "吸收存款及同业存放",
        "利息净收入",
        "手续费及佣金净收入",
        "发行债务证券所收到的现金",
    }
    if "银行" in joined_titles or markers & bank_markers:
        return "bank"
    return "company"


def _build_display_name(path: Path, kind: str) -> str:
    prefix = "银行财务" if kind == "bank" else "公司财报"
    return f"{prefix} - {path.stem}"


@lru_cache(maxsize=1)
def discover_templates() -> tuple[TemplateSpec, ...]:
    project_root = resolve_project_root()
    candidates: dict[str, Path] = {}
    for pattern in TEMPLATE_GLOB_PATTERNS:
        for path in sorted(project_root.glob(pattern)):
            candidates.setdefault(path.stem, path)

    templates: list[TemplateSpec] = []
    for template_id, path in sorted(candidates.items(), key=lambda item: item[0]):
        kind = _guess_template_kind(path)
        templates.append(
            TemplateSpec(
                template_id=template_id,
                display_name=_build_display_name(path, kind),
                kind=kind,
                path=path,
            )
        )
    return tuple(templates)


def available_template_ids() -> tuple[str, ...]:
    return tuple(template.template_id for template in discover_templates())


def default_template_id() -> str | None:
    templates = discover_templates()
    if not templates:
        return None
    for template in templates:
        if template.kind == "company":
            return template.template_id
    return templates[0].template_id


def resolve_template(template_id: str | None) -> TemplateSpec:
    templates = discover_templates()
    if not templates:
        raise FileNotFoundError("未找到可用的 Excel 模板，请把模板文件放到 cninfo_pipeline 目录。")

    requested = str(template_id or default_template_id() or "").strip()
    if requested:
        for template in templates:
            if template.template_id == requested:
                return template

    fallback_id = default_template_id()
    if fallback_id:
        for template in templates:
            if template.template_id == fallback_id:
                return template
    return templates[0]
