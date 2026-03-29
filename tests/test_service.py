from __future__ import annotations

import shutil
import unittest
from pathlib import Path

from openpyxl import load_workbook

from cninfo_pipeline.client import CompanyRecord
from cninfo_pipeline.service import (
    build_statement_matrix,
    export_statement_workbook,
    filter_annual_merged_records,
    prepare_cache_dir,
)


class ServiceTests(unittest.TestCase):
    def test_filter_annual_merged_records(self) -> None:
        records = [
            {"ENDDATE": "2024-12-31", "F003V": "合并本期", "F010N": 1},
            {"ENDDATE": "2024-09-30", "F003V": "合并本期", "F010N": 2},
            {"ENDDATE": "2024-12-31", "F003V": "母公司本期", "F010N": 3},
            {"ENDDATE": "2023-12-31", "F003V": "合并本期", "F010N": 4},
        ]
        filtered = filter_annual_merged_records(records)
        self.assertEqual([item["ENDDATE"] for item in filtered], ["2024-12-31", "2023-12-31"])

    def test_build_statement_matrix(self) -> None:
        records = [
            {
                "ENDDATE": "2024-12-31",
                "F006N": 100.0,
                "F008N": 10.0,
                "F009N": 20.0,
                "F010N": 30.0,
                "F011N": 40.0,
                "F015N": 50.0,
                "F018N": 60.0,
                "F019N": 70.0,
                "F025N": 80.0,
                "F026N": 90.0,
                "F033N": 5.0,
                "F037N": 1000.0,
                "F038N": 1070.0,
                "F041N": 11.0,
                "F042N": 22.0,
                "F048N": 44.0,
                "F047N": 4.0,
                "F052N": 300.0,
                "F060N": 200.0,
                "F061N": 500.0,
                "F062N": 200.0,
                "F063N": 100.0,
                "F064N": 50.0,
                "F065N": 150.0,
                "F067N": 20.0,
                "F070N": 570.0,
                "F073N": 550.0,
                "F074N": 7.0,
                "F121N": 9.0,
                "F122N": 8.0,
            }
        ]

        matrix = build_statement_matrix(records)
        rows = {row[0]: row[1:] for row in matrix}

        self.assertEqual(rows["报表日期"], [20241231])
        self.assertEqual(rows["单位"], ["元"])
        self.assertEqual(rows["货币资金"], [100.0])
        self.assertEqual(rows["应收票据及应收账款8+9"], [30.0])
        self.assertEqual(rows["在建工程(合计)32+33"], [90.0])
        self.assertEqual(rows["固定资产及清理(合计)35+36"], [80.0])
        self.assertEqual(rows["应付票据及应付账款53+54"], [33.0])
        self.assertEqual(rows["其他应付款"], [40.0])
        self.assertAlmostEqual(rows["实际资产负债率"][0], 500.0 / 1070.0)
        self.assertAlmostEqual(rows["FA&CIP占比资产总额"][0], 170.0 / 1070.0)
        self.assertAlmostEqual(rows["商誉占比归属股东权益"][0], 5.0 / 550.0)
        self.assertEqual(rows["FA&CIP净额"], [170.0])
        self.assertEqual(rows["FA&CIP净额-合计1"], [80.0])
        self.assertEqual(rows["FA&CIP净额-合计2"], [170.0])
        self.assertAlmostEqual(rows["权益乘数=资产总额/所有者权益"][0], 1070.0 / 570.0)
        self.assertAlmostEqual(rows["产权比率=负债总额/归属股东权益"][0], 500.0 / 550.0)

    def test_build_statement_matrix_supports_unit_scaling(self) -> None:
        records = [
            {
                "ENDDATE": "2024-12-31",
                "F006N": 2_500_000.0,
                "F025N": 6_000_000.0,
                "F026N": 4_000_000.0,
                "F038N": 20_000_000.0,
                "F061N": 5_000_000.0,
                "F070N": 15_000_000.0,
                "F073N": 10_000_000.0,
            }
        ]

        matrix = build_statement_matrix(records, unit_label="万元")
        rows = {row[0]: row[1:] for row in matrix}

        self.assertEqual(rows["单位"], ["万元"])
        self.assertEqual(rows["货币资金"], [250.0])
        self.assertEqual(rows["固定资产及清理(合计)35+36"], [600.0])
        self.assertEqual(rows["FA&CIP净额"], [1000.0])
        self.assertEqual(rows["资产总计"], [2000.0])
        self.assertAlmostEqual(rows["实际资产负债率"][0], 0.25)
        self.assertAlmostEqual(rows["权益乘数=资产总额/所有者权益"][0], 20_000_000.0 / 15_000_000.0)

    def test_export_statement_workbook(self) -> None:
        company = CompanyRecord(seccode="600900", secname="长江电力", orgname="中国长江电力股份有限公司")
        matrix = [
            ["报表日期", 20241231, 20231231],
            ["单位", "元", "元"],
            ["流动资产", None, None],
            ["货币资金", 100.0, 90.0],
        ]

        temp_dir = Path("test_artifacts")
        shutil.rmtree(temp_dir, ignore_errors=True)
        temp_dir.mkdir(parents=True, exist_ok=True)
        try:
            workbook_path = export_statement_workbook(company, matrix, temp_dir)
            self.assertTrue(workbook_path.exists())

            workbook = load_workbook(workbook_path, data_only=True)
            sheet = workbook.active
            self.assertEqual(sheet["A1"].value, "报表日期")
            self.assertEqual(sheet["B1"].value, 20241231)
            self.assertEqual(sheet["B2"].value, "元")
            self.assertEqual(sheet["A4"].value, "货币资金")
            self.assertEqual(sheet["C4"].value, 90.0)
            self.assertEqual(sheet["A1"].font.name, "等线")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_prepare_cache_dir_uses_hidden_internal_dir(self) -> None:
        temp_dir = Path("test_artifacts")
        shutil.rmtree(temp_dir, ignore_errors=True)
        temp_dir.mkdir(parents=True, exist_ok=True)
        try:
            legacy_cache = temp_dir / ".cache"
            legacy_cache.mkdir(parents=True, exist_ok=True)
            (legacy_cache / "companies.json").write_text("{}", encoding="utf-8")

            cache_dir = prepare_cache_dir(temp_dir)

            self.assertEqual(cache_dir.name, ".cninfo_internal")
            self.assertTrue(cache_dir.exists())
            self.assertFalse(legacy_cache.exists())
            self.assertTrue((cache_dir / "companies.json").exists())
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
