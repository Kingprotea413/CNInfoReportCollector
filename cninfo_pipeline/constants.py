from __future__ import annotations


UNIT_SCALE_MAP = {
    "元": 1,
    "千元": 1_000,
    "万元": 10_000,
    "亿元": 100_000_000,
}
DEFAULT_UNIT_LABEL = "元"
AVAILABLE_UNIT_LABELS = tuple(UNIT_SCALE_MAP.keys())
