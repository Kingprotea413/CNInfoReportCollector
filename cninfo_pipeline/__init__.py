"""CNInfo annual balance-sheet collection pipeline."""

from __future__ import annotations

from typing import TYPE_CHECKING

from .constants import AVAILABLE_UNIT_LABELS, DEFAULT_UNIT_LABEL

if TYPE_CHECKING:
    from .service import AnnualReportPipeline, PipelineResult

__all__ = [
    "AnnualReportPipeline",
    "AVAILABLE_UNIT_LABELS",
    "DEFAULT_UNIT_LABEL",
    "PipelineResult",
]


def __getattr__(name: str):
    if name in {"AnnualReportPipeline", "PipelineResult"}:
        from .service import AnnualReportPipeline, PipelineResult

        exported = {
            "AnnualReportPipeline": AnnualReportPipeline,
            "PipelineResult": PipelineResult,
        }
        return exported[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
