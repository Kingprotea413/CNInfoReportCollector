from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook, load_workbook

from cninfo_pipeline.client import CompanyRecord
from cninfo_pipeline.service import (
    AnnualReportPipeline,
    build_statement_matrix,
    export_statement_workbook,
    filter_annual_merged_records,
    prepare_cache_dir,
)
from cninfo_pipeline.template_export import (
    MISSING_REASON_SHEET,
    SECTION_FILL_BY_STATEMENT,
    export_template_workbook,
    is_section_label,
)
from cninfo_pipeline.template_registry import TemplateSpec, discover_templates, resolve_template


def build_balance_record(date: str, **fields: float) -> dict:
    return {
        "ENDDATE": date,
        "F001D": date,
        "F002V": "071001",
        "F003V": "合并本期",
        "F005V": "定期报告",
        **fields,
    }


def build_statement_record(date: str, **fields: float) -> dict:
    return {
        "ENDDATE": date,
        "F001D": date,
        "F002V": "071001",
        "F003V": "合并本期",
        "F005V": "定期报告",
        **fields,
    }


def find_row_index(sheet, label: str) -> int:
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row, 1).value == label:
            return row
    raise AssertionError(f"未找到标签：{label}")


def iso_date(value: object) -> str:
    if hasattr(value, "date"):
        return value.date().isoformat()
    if hasattr(value, "isoformat"):
        return value.isoformat()
    raise AssertionError(f"不是日期值：{value!r}")


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

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = export_statement_workbook(company, matrix, Path(tmpdir))
            self.assertTrue(workbook_path.exists())

            workbook = load_workbook(workbook_path, data_only=True)
            sheet = workbook.active
            self.assertEqual(sheet["A1"].value, "报表日期")
            self.assertEqual(sheet["B1"].value, 20241231)
            self.assertEqual(sheet["B2"].value, "元")
            self.assertEqual(sheet["A4"].value, "货币资金")
            self.assertEqual(sheet["C4"].value, 90.0)
            self.assertEqual(sheet["A1"].font.name, "等线")

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

    def test_prepare_cache_dir_defaults_to_app_cache_dir(self) -> None:
        temp_dir = Path("test_artifacts")
        shutil.rmtree(temp_dir, ignore_errors=True)
        try:
            expected = temp_dir / "app-cache"
            with patch("cninfo_pipeline.service.resolve_default_cache_dir", return_value=expected):
                cache_dir = prepare_cache_dir()

            self.assertEqual(cache_dir, expected)
            self.assertTrue(cache_dir.exists())
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_discover_templates_finds_bank_and_company(self) -> None:
        templates = {template.template_id: template for template in discover_templates()}

        self.assertIn("银行财务报表模版", templates)
        self.assertIn("公司财务报表模版", templates)
        self.assertEqual(templates["公司财务报表模版"].kind, "company")
        self.assertEqual(resolve_template("工商银行财务报表模版").template_id, "银行财务报表模版")
        self.assertEqual(resolve_template("长江电力年度报告财务报表模版").template_id, "公司财务报表模版")

    def test_is_section_label_does_not_treat_parenthesized_items_as_titles(self) -> None:
        self.assertFalse(is_section_label("（3）现金流量套期储备", None))
        self.assertFalse(is_section_label("(2)其他权益工具投资公允价值变动", None))
        self.assertTrue(is_section_label("一、经营活动产生的现金流量：", None))
        self.assertTrue(is_section_label("资产：", None))

    def test_pipeline_run_uses_selected_template_export(self) -> None:
        company = CompanyRecord(seccode="601138", secname="工业富联", orgname="富士康工业互联网股份有限公司")
        balance_records = [build_balance_record("2024-12-31"), build_balance_record("2023-12-31")]
        income_records = [build_statement_record("2024-12-31"), build_statement_record("2023-12-31")]
        cash_flow_records = [build_statement_record("2024-12-31"), build_statement_record("2023-12-31")]
        client = MagicMock()
        client.search_company.return_value = company
        client.fetch_balance_sheet.return_value = balance_records
        client.fetch_income_statement.return_value = income_records
        client.fetch_cash_flow_statement.return_value = cash_flow_records

        pipeline = AnnualReportPipeline(client=client)
        fake_output = Path("out") / "工业富联财务报表2024YE.xlsx"
        fake_cache_dir = Path("cache-dir")
        fake_output_dir = Path("exports")
        template_id = "银行财务报表模版"

        with (
            patch("cninfo_pipeline.service.prepare_cache_dir", return_value=fake_cache_dir),
            patch("cninfo_pipeline.service.resolve_default_output_dir", return_value=fake_output_dir),
            patch("cninfo_pipeline.template_export.export_template_workbook", return_value=fake_output) as exporter,
        ):
            result = pipeline.run(company_query="工业富联", unit_label="万元", template_id=template_id)

        client.set_cache_dir.assert_called_once_with(fake_cache_dir)
        client.search_company.assert_called_once_with("工业富联")
        client.fetch_balance_sheet.assert_called_once_with("601138")
        client.fetch_income_statement.assert_called_once_with("601138")
        client.fetch_cash_flow_statement.assert_called_once_with("601138")
        exporter.assert_called_once_with(
            company=company,
            balance_records=balance_records,
            income_records=income_records,
            cash_flow_records=cash_flow_records,
            output_dir=fake_output_dir,
            unit_label="万元",
            template_id=template_id,
        )
        self.assertEqual(result.output_path, fake_output)
        self.assertEqual(result.total_records, 2)
        self.assertEqual(result.annual_records, 2)
        self.assertEqual(result.unit_label, "万元")
        self.assertEqual(result.template_id, template_id)
        self.assertEqual(result.template_name, resolve_template(template_id).display_name)

    def test_export_template_workbook_company_template_uses_template_sheets(self) -> None:
        company = CompanyRecord(seccode="600900", secname="长江电力", orgname="中国长江电力股份有限公司")
        template = resolve_template("公司财务报表模版")
        balance_records = [
            build_balance_record("2024-12-31", F006N=10_000_000, F038N=28_000_000, F061N=15_000_000),
            build_balance_record("2023-12-31", F006N=8_000_000, F038N=24_000_000, F061N=12_000_000),
        ]
        income_records = [
            build_statement_record("2024-12-31", F006N=30_000_000, F035N=30_000_000, F018N=10_000_000),
            build_statement_record("2023-12-31", F006N=25_000_000, F035N=25_000_000, F018N=8_000_000),
        ]
        cash_flow_records = [
            build_statement_record("2024-12-31", F015N=5_000_000, F041N=9_000_000),
            build_statement_record("2023-12-31", F015N=4_000_000, F041N=8_000_000),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = export_template_workbook(
                company=company,
                balance_records=balance_records,
                income_records=income_records,
                cash_flow_records=cash_flow_records,
                output_dir=tmpdir,
                unit_label="万元",
                template_id=template.template_id,
            )

            workbook = load_workbook(workbook_path, data_only=False)
            template_workbook = load_workbook(template.path, data_only=False)
            self.assertEqual(workbook.sheetnames[:-1], template_workbook.sheetnames)
            self.assertEqual(workbook.sheetnames[-1], MISSING_REASON_SHEET)

            balance_sheet = workbook[template_workbook.sheetnames[0]]
            self.assertEqual(balance_sheet["A1"].value, "项目（单位：万元）")
            self.assertEqual(iso_date(balance_sheet["C1"].value), "2024-12-31")
            self.assertEqual(iso_date(balance_sheet["D1"].value), "2023-12-31")
            self.assertGreater(balance_sheet.column_dimensions["A"].width, 10)
            self.assertTrue(balance_sheet["A2"].alignment.wrap_text)
            self.assertEqual(balance_sheet["A1"].font.name, template_workbook[template_workbook.sheetnames[0]]["A1"].font.name)
            self.assertEqual(balance_sheet["A1"].font.sz, template_workbook[template_workbook.sheetnames[0]]["A1"].font.sz)
            self.assertEqual(balance_sheet["C1"].font.sz, template_workbook[template_workbook.sheetnames[0]]["A1"].font.sz)
            self.assertEqual(balance_sheet["A1"].border.left.style, "medium")
            self.assertEqual(balance_sheet["C1"].border.left.style, "medium")
            cash_row = find_row_index(balance_sheet, "货币资金")
            self.assertEqual(balance_sheet.cell(cash_row, 3).value, 1000)
            self.assertEqual(balance_sheet.cell(cash_row, 4).value, 800)

            income_sheet = workbook[template_workbook.sheetnames[1]]
            self.assertEqual(income_sheet["A1"].value, "项目（单位：万元）")
            self.assertEqual(iso_date(income_sheet["B1"].value), "2024-12-31")
            self.assertEqual(iso_date(income_sheet["C1"].value), "2023-12-31")

            cash_sheet = workbook[template_workbook.sheetnames[2]]
            self.assertEqual(cash_sheet["A1"].value, "项目（单位：万元）")
            self.assertEqual(iso_date(cash_sheet["B1"].value), "2024-12-31")
            self.assertEqual(iso_date(cash_sheet["C1"].value), "2023-12-31")

            template_workbook.close()
            workbook.close()

    def test_export_template_workbook_expands_to_all_annual_periods(self) -> None:
        company = CompanyRecord(seccode="600900", secname="长江电力", orgname="中国长江电力股份有限公司")
        template = resolve_template("公司财务报表模版")
        balance_records = [
            build_balance_record("2024-12-31", F006N=10_000_000, F038N=28_000_000, F061N=15_000_000),
            build_balance_record("2023-12-31", F006N=8_000_000, F038N=24_000_000, F061N=12_000_000),
            build_balance_record("2022-12-31", F006N=6_000_000, F038N=20_000_000, F061N=10_000_000),
        ]
        income_records = [
            build_statement_record("2024-12-31", F006N=30_000_000, F035N=30_000_000, F018N=10_000_000),
            build_statement_record("2023-12-31", F006N=25_000_000, F035N=25_000_000, F018N=8_000_000),
            build_statement_record("2022-12-31", F006N=20_000_000, F035N=20_000_000, F018N=6_000_000),
        ]
        cash_flow_records = [
            build_statement_record("2024-12-31", F015N=5_000_000, F041N=9_000_000),
            build_statement_record("2023-12-31", F015N=4_000_000, F041N=8_000_000),
            build_statement_record("2022-12-31", F015N=3_000_000, F041N=7_000_000),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = export_template_workbook(
                company=company,
                balance_records=balance_records,
                income_records=income_records,
                cash_flow_records=cash_flow_records,
                output_dir=tmpdir,
                unit_label="万元",
                template_id=template.template_id,
            )

            workbook = load_workbook(workbook_path, data_only=False)
            balance_sheet = workbook[workbook.sheetnames[0]]
            self.assertEqual(iso_date(balance_sheet["C1"].value), "2024-12-31")
            self.assertEqual(iso_date(balance_sheet["D1"].value), "2023-12-31")
            self.assertEqual(iso_date(balance_sheet["E1"].value), "2022-12-31")
            cash_row = find_row_index(balance_sheet, "货币资金")
            self.assertEqual(balance_sheet.cell(cash_row, 5).value, 600)
            workbook.close()

    def test_export_template_workbook_bank_template_inserts_year_headers(self) -> None:
        company = CompanyRecord(seccode="601398", secname="工商银行", orgname="中国工商银行股份有限公司")
        balance_records = [
            build_balance_record(
                "2025-12-31",
                F023N=3_000_000,
                F061N=16_000_000,
                F062N=500_000,
                F063N=600_000,
                F064N=700_000,
                F065N=800_000,
                F067N=90_000,
                F070N=5_000_000,
                F071N=21_000_000,
                F073N=4_910_000,
                F074N=110_000,
                F076N=900_000,
                F089N=1_200_000,
                F090N=1_300_000,
                F103N=1_400_000,
                F104N=500_000,
                F105N=900_000,
            ),
            build_balance_record(
                "2024-12-31",
                F023N=2_500_000,
                F061N=14_000_000,
                F062N=450_000,
                F063N=550_000,
                F064N=650_000,
                F065N=750_000,
                F067N=80_000,
                F070N=4_500_000,
                F071N=18_500_000,
                F073N=4_420_000,
                F074N=90_000,
                F076N=850_000,
                F089N=1_000_000,
                F090N=1_100_000,
                F103N=1_200_000,
                F104N=400_000,
                F105N=800_000,
            ),
        ]
        income_records = [
            build_statement_record("2025-12-31", F033N=500_000, F035N=900_000, F027N=300_000),
            build_statement_record("2024-12-31", F033N=400_000, F035N=800_000, F027N=250_000),
        ]
        cash_flow_records = [
            build_statement_record(
                "2025-12-31",
                F009N=700_000,
                F014N=600_000,
                F015N=100_000,
                F011N=40_000,
                F012N=20_000,
                F020N=30_000,
                F022N=50_000,
                F031N=90_000,
                F032N=80_000,
                F034N=70_000,
                F037N=60_000,
                F040N=500_000,
                F041N=400_000,
                F076N=110_000,
                F081N=120_000,
                F084N=130_000,
                F087N=140_000,
            ),
            build_statement_record(
                "2024-12-31",
                F009N=650_000,
                F014N=580_000,
                F015N=70_000,
                F011N=35_000,
                F012N=15_000,
                F020N=25_000,
                F022N=45_000,
                F031N=85_000,
                F032N=75_000,
                F034N=65_000,
                F037N=55_000,
                F040N=450_000,
                F041N=350_000,
                F076N=100_000,
                F081N=110_000,
                F084N=120_000,
                F087N=130_000,
            ),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "bank_template.xlsx"
            template_workbook = Workbook()

            balance_sheet = template_workbook.active
            balance_sheet.title = "测试银行资产负债表"
            for row, label in enumerate(
                [
                    "资产：",
                    "长期股权投资",
                    "拆入资金",
                    "衍生金融负债",
                    "其他权益工具：",
                    "优先股",
                    "永续债",
                    "资本公积",
                    "其他综合收益",
                    "盈余公积",
                    "一般准备",
                    "未分配利润",
                    "少数股东权益",
                    "负债合计",
                ],
                start=1,
            ):
                balance_sheet.cell(row, 1, label)

            income_sheet = template_workbook.create_sheet("测试银行利润表")
            for row, label in enumerate(["利息净收入", "营业收入"], start=1):
                income_sheet.cell(row, 1, label)

            cash_sheet = template_workbook.create_sheet("测试银行现金流量表")
            for row, label in enumerate(
                [
                    "一、经营活动产生的现金流量：",
                    "收取的利息、手续费及佣金的现金",
                    "客户贷款及垫款净额",
                    "支付给职工以及为职工支付的现金支付的各项税费",
                    "收到其他与投资活动有关的现金",
                    "购建固定资产、无形资产和其他长期资产支付的现金",
                    "发行债务证券所收到的现金",
                    "偿还债务证券所支付的现金",
                    "支付其他与筹资活动有关的现金",
                    "四、汇率变动对现金及现金等价物",
                    "经营活动产生的现金流量净额",
                    "年初现金及现金等价物余额",
                    "年末现金及现金等价物余额",
                ],
                start=1,
            ):
                cash_sheet.cell(row, 1, label)

            template_workbook.save(template_path)
            template_workbook.close()

            template = TemplateSpec(
                template_id="bank-test",
                display_name="银行财务 - 测试模板",
                kind="bank",
                path=template_path,
            )

            with patch("cninfo_pipeline.template_export.resolve_template", return_value=template):
                workbook_path = export_template_workbook(
                    company=company,
                    balance_records=balance_records,
                    income_records=income_records,
                    cash_flow_records=cash_flow_records,
                    output_dir=tmpdir,
                    unit_label="万元",
                    template_id=template.template_id,
                )

            workbook = load_workbook(workbook_path, data_only=False)
            balance_sheet = workbook["测试银行资产负债表"]
            income_sheet = workbook["测试银行利润表"]
            cash_sheet = workbook["测试银行现金流量表"]
            reason_sheet = workbook[MISSING_REASON_SHEET]

            self.assertEqual(balance_sheet["A1"].value, "项目（单位：万元）")
            self.assertEqual(iso_date(balance_sheet["B1"].value), "2025-12-31")
            self.assertEqual(iso_date(balance_sheet["C1"].value), "2024-12-31")
            self.assertEqual(balance_sheet["A2"].value, "资产：")
            self.assertGreater(balance_sheet.column_dimensions["A"].width, 10)
            self.assertTrue(balance_sheet["A2"].alignment.wrap_text)
            self.assertTrue(balance_sheet["A2"].fill.fgColor.rgb.endswith(SECTION_FILL_BY_STATEMENT["balance"].fgColor.rgb[-6:]))
            self.assertIsNone(balance_sheet["A2"].comment)
            self.assertEqual(balance_sheet["A1"].font.sz, 11)
            self.assertEqual(balance_sheet["B1"].font.sz, 11)
            self.assertEqual(balance_sheet["A1"].border.left.style, "medium")
            self.assertEqual(balance_sheet["B1"].border.left.style, "medium")
            long_term_row = find_row_index(balance_sheet, "长期股权投资")
            self.assertEqual(balance_sheet.cell(long_term_row, 2).value, 300)
            split_borrow_row = find_row_index(balance_sheet, "拆入资金")
            self.assertEqual(balance_sheet.cell(split_borrow_row, 2).value, 120)
            derivative_liab_row = find_row_index(balance_sheet, "衍生金融负债")
            self.assertEqual(balance_sheet.cell(derivative_liab_row, 2).value, 130)
            equity_tool_row = find_row_index(balance_sheet, "其他权益工具：")
            self.assertEqual(balance_sheet.cell(equity_tool_row, 2).value, 140)
            self.assertNotEqual(
                balance_sheet.cell(equity_tool_row, 1).fill.fgColor.rgb,
                SECTION_FILL_BY_STATEMENT["balance"].fgColor.rgb,
            )
            preferred_row = find_row_index(balance_sheet, "优先股")
            self.assertEqual(balance_sheet.cell(preferred_row, 2).value, 50)
            perpetual_row = find_row_index(balance_sheet, "永续债")
            self.assertEqual(balance_sheet.cell(perpetual_row, 2).value, 90)
            oci_row = find_row_index(balance_sheet, "其他综合收益")
            self.assertEqual(balance_sheet.cell(oci_row, 2).value, 11)
            minority_row = find_row_index(balance_sheet, "少数股东权益")
            self.assertEqual(balance_sheet.cell(minority_row, 2).value, 9)

            self.assertEqual(income_sheet["A2"].value, "利息净收入")
            self.assertEqual(iso_date(income_sheet["B1"].value), "2025-12-31")
            interest_row = find_row_index(income_sheet, "利息净收入")
            self.assertEqual(income_sheet.cell(interest_row, 2).value, 50)

            self.assertEqual(cash_sheet["A2"].value, "一、经营活动产生的现金流量：")
            self.assertEqual(iso_date(cash_sheet["B1"].value), "2025-12-31")
            self.assertTrue(cash_sheet["A2"].fill.fgColor.rgb.endswith(SECTION_FILL_BY_STATEMENT["cash"].fgColor.rgb[-6:]))
            collected_fee_row = find_row_index(cash_sheet, "收取的利息、手续费及佣金的现金")
            self.assertEqual(cash_sheet.cell(collected_fee_row, 2).value, 12)
            loans_row = find_row_index(cash_sheet, "客户贷款及垫款净额")
            self.assertEqual(cash_sheet.cell(loans_row, 2).value, 13)
            merged_payroll_tax_row = find_row_index(cash_sheet, "支付给职工以及为职工支付的现金支付的各项税费")
            self.assertEqual(cash_sheet.cell(merged_payroll_tax_row, 2).value, 6)
            other_invest_row = find_row_index(cash_sheet, "收到其他与投资活动有关的现金")
            self.assertEqual(cash_sheet.cell(other_invest_row, 2).value, 3)
            capex_row = find_row_index(cash_sheet, "购建固定资产、无形资产和其他长期资产支付的现金")
            self.assertEqual(cash_sheet.cell(capex_row, 2).value, 5)
            debt_issue_row = find_row_index(cash_sheet, "发行债务证券所收到的现金")
            self.assertEqual(cash_sheet.cell(debt_issue_row, 2).value, 11)
            debt_repay_row = find_row_index(cash_sheet, "偿还债务证券所支付的现金")
            self.assertEqual(cash_sheet.cell(debt_repay_row, 2).value, 8)
            other_finance_row = find_row_index(cash_sheet, "支付其他与筹资活动有关的现金")
            self.assertEqual(cash_sheet.cell(other_finance_row, 2).value, 7)
            fx_row = find_row_index(cash_sheet, "四、汇率变动对现金及现金等价物")
            self.assertEqual(cash_sheet.cell(fx_row, 2).value, 6)
            net_cash_row = find_row_index(cash_sheet, "经营活动产生的现金流量净额")
            self.assertEqual(cash_sheet.cell(net_cash_row, 2).value, 10)
            opening_cash_row = find_row_index(cash_sheet, "年初现金及现金等价物余额")
            self.assertEqual(cash_sheet.cell(opening_cash_row, 2).value, 50)
            closing_cash_row = find_row_index(cash_sheet, "年末现金及现金等价物余额")
            self.assertEqual(cash_sheet.cell(closing_cash_row, 2).value, 40)
            self.assertEqual(reason_sheet["A1"].value, "工作表")
            reason_labels = {reason_sheet.cell(row, 3).value for row in range(2, reason_sheet.max_row + 1)}
            self.assertIn("资产：", reason_labels)
            self.assertIn("一、经营活动产生的现金流量：", reason_labels)

            workbook.close()

    def test_export_template_workbook_explains_missing_rows(self) -> None:
        company = CompanyRecord(seccode="601398", secname="工商银行", orgname="中国工商银行股份有限公司")
        balance_records = [
            build_balance_record("2025-12-31", F063N=None),
            build_balance_record("2024-12-31", F063N=None),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "bank_missing_template.xlsx"
            template_workbook = Workbook()
            sheet = template_workbook.active
            sheet.title = "测试银行资产负债表"
            for row, label in enumerate(["资产：", "未映射项目", "优先股", "资本公积"], start=1):
                sheet.cell(row, 1, label)
            template_workbook.save(template_path)
            template_workbook.close()

            template = TemplateSpec(
                template_id="bank-missing-test",
                display_name="银行财务 - 缺口测试模板",
                kind="bank",
                path=template_path,
            )

            with patch("cninfo_pipeline.template_export.resolve_template", return_value=template):
                workbook_path = export_template_workbook(
                    company=company,
                    balance_records=balance_records,
                    income_records=[],
                    cash_flow_records=[],
                    output_dir=tmpdir,
                    unit_label="万元",
                    template_id=template.template_id,
                )

            workbook = load_workbook(workbook_path, data_only=False)
            reason_sheet = workbook[MISSING_REASON_SHEET]
            reasons = {
                reason_sheet.cell(row, 3).value: reason_sheet.cell(row, 4).value
                for row in range(2, reason_sheet.max_row + 1)
            }

            self.assertEqual(reasons["资产："], "模板分组行")
            self.assertEqual(reasons["未映射项目"], "待补充映射/公式")
            self.assertEqual(reasons["优先股"], "API 未提供字段")
            self.assertEqual(reasons["资本公积"], "本期未披露数值")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
