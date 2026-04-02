from __future__ import annotations

from copy import copy
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .client import CompanyRecord
from .constants import DEFAULT_UNIT_LABEL, UNIT_SCALE_MAP
from .paths import ensure_writable_dir
from .service import normalize_unit_label
from .template_registry import TemplateSpec, resolve_template


Resolver = Callable[[dict], object | None]
DEFAULT_ANNUAL_PERIODS = 2
DATE_PATTERN = re.compile(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
HEADER_FONT = Font(bold=True, color="1F2937")
SECTION_FILL_BY_STATEMENT = {
    "balance": PatternFill(fill_type="solid", fgColor="DCEBFA"),
    "income": PatternFill(fill_type="solid", fgColor="E3F4E8"),
    "cash": PatternFill(fill_type="solid", fgColor="FCE8D5"),
}
SECTION_FONT = Font(bold=True, color="1F2937")
MISSING_REASON_SHEET = "空白项说明"
SEPARATOR_SIDE = Side(style="medium", color="000000")
STATEMENT_SHEET_SUFFIX = {
    "balance": "资产负债表",
    "income": "利润表",
    "cash": "现金流量表",
}
SECTION_LABEL_EXACT = {
    "资产",
    "负债",
    "所有者权益",
    "股东权益",
    "流动资产",
    "非流动资产",
    "流动负债",
    "非流动负债",
    "经营活动产生的现金流量",
    "投资活动产生的现金流量",
    "筹资活动产生的现金流量",
    "现金及现金等价物净增加额",
    "补充资料",
}


def _resolver_with_metadata(
    resolver: Resolver,
    *,
    kind: str,
    source_fields: tuple[str, ...] = (),
) -> Resolver:
    setattr(resolver, "_resolver_kind", kind)
    setattr(resolver, "_source_fields", tuple(dict.fromkeys(source_fields)))
    return resolver


def resolver_kind(resolver: Resolver | None) -> str:
    return str(getattr(resolver, "_resolver_kind", "custom"))


def resolver_source_fields(resolver: Resolver | None) -> tuple[str, ...]:
    return tuple(getattr(resolver, "_source_fields", ()))


def field(field_name: str) -> Resolver:
    return _resolver_with_metadata(
        lambda record: record.get(field_name),
        kind="field",
        source_fields=(field_name,),
    )


def sum_fields(*field_names: str) -> Resolver:
    def resolver(record: dict) -> object | None:
        values = [record.get(field_name) for field_name in field_names]
        numeric = [value for value in values if isinstance(value, (int, float))]
        return sum(numeric) if numeric else None

    return _resolver_with_metadata(
        resolver,
        kind="sum",
        source_fields=tuple(field_names),
    )


def subtract_fields(total_field: str, *part_fields: str) -> Resolver:
    def resolver(record: dict) -> object | None:
        total_value = record.get(total_field)
        if not isinstance(total_value, (int, float)):
            return None
        parts = [record.get(field_name) for field_name in part_fields]
        numeric_parts = [value for value in parts if isinstance(value, (int, float))]
        return total_value - sum(numeric_parts)

    return _resolver_with_metadata(
        resolver,
        kind="subtract",
        source_fields=(total_field, *part_fields),
    )


def first_available(*resolvers: Resolver) -> Resolver:
    def resolver(record: dict) -> object | None:
        for candidate in resolvers:
            value = candidate(record)
            if value is not None:
                return value
        return None

    source_fields: list[str] = []
    for candidate in resolvers:
        source_fields.extend(resolver_source_fields(candidate))
    return _resolver_with_metadata(
        resolver,
        kind="first_available",
        source_fields=tuple(source_fields),
    )


def normalize_label(value: object | None) -> str:
    text = str(value or "").strip()
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
        "，": "",
        ",": "",
        "。": "",
        "“": "",
        "”": "",
        "‘": "",
        "’": "",
        "－": "",
        "-": "",
        "/": "",
        "／": "",
        "·": "",
        "\n": "",
        "\r": "",
        "\t": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


LABEL_FIXUPS = {
    normalize_label("长期待滩费用"): normalize_label("长期待摊费用"),
    normalize_label("其中：营业收入"): normalize_label("营业收入"),
    normalize_label("其中：营业成本"): normalize_label("营业成本"),
    normalize_label("其中：应收利息"): normalize_label("应收利息"),
    normalize_label("其中：应付利息"): normalize_label("应付利息"),
    normalize_label("实收资本（或股本）"): normalize_label("实收资本(或股本)"),
    normalize_label("所有者权益（或股东权益）"): normalize_label("所有者权益"),
    normalize_label("归属于母公司所有者权益 （或股东权益）合计"): normalize_label("归属于母公司所有者权益（或股东权益）合计"),
    normalize_label("所有者权益（或股东权益）合计"): normalize_label("所有者权益(或股东权益)合计"),
    normalize_label("负债和所有者权益（或股东权益）总计"): normalize_label("负债和所有者权益(或股东权益)总计"),
    normalize_label("经营活动产生的现金量"): normalize_label("经营活动产生的现金流量"),
    normalize_label("经营活动产生的现金量："): normalize_label("经营活动产生的现金流量："),
    normalize_label("客户存款和同业存放款项净增加額"): normalize_label("客户存款和同业存放款项净增加额"),
    normalize_label("向其他金融机构拆入资金净增加額"): normalize_label("向其他金融机构拆入资金净增加额"),
    normalize_label("收取利息、手续费及佣金的现"): normalize_label("收取利息、手续费及佣金的现金"),
    normalize_label("取得惜款收到的现金"): normalize_label("取得借款收到的现金"),
    normalize_label("收到其他与等资活动有关的现金"): normalize_label("收到其他与筹资活动有关的现金"),
    normalize_label("支付其他与等资活动有关的现金"): normalize_label("支付其他与筹资活动有关的现金"),
    normalize_label("投资动产生的现金流量净额"): normalize_label("投资活动产生的现金流量净额"),
    normalize_label("五、现金及现金等价物净增加额加期初现金及现金等价物余额"): normalize_label("现金及现金等价物净增加额"),
    normalize_label("六、期末现金及现金等价物余额"): normalize_label("期末现金及现金等价物余额"),
    normalize_label("五、现金及现金等价物净变动额加：年初现金及现金等价物余额"): normalize_label("现金及现金等价物净变动额"),
    normalize_label("六、年末现金及现金等价物余额"): normalize_label("年末现金及现金等价物余额"),
    normalize_label("四、汇率变动对现金及现金等价物"): normalize_label("四、汇率变动对现金及现金等价物的影响"),
}


def canonical_label(value: object | None) -> str:
    normalized = normalize_label(value)
    return LABEL_FIXUPS.get(normalized, normalized)


NON_SCALED_LABELS = {
    canonical_label("基本每股收益（元/股）"),
    canonical_label("稀释每股收益（元/股）"),
    canonical_label("基本每股收益（人民币元）"),
    canonical_label("稀释每股收益（人民币元）"),
}


COMPANY_BALANCE_RESOLVERS: dict[str, Resolver] = {
    canonical_label("货币资金"): field("F006N"),
    canonical_label("交易性金融资产"): field("F117N"),
    canonical_label("衍生金融资产"): field("F080N"),
    canonical_label("应收票据"): field("F008N"),
    canonical_label("应收账款"): field("F009N"),
    canonical_label("应收款项融资"): field("F110N"),
    canonical_label("预付款项"): field("F010N"),
    canonical_label("其他应收款"): field("F011N"),
    canonical_label("应收利息"): field("F012N"),
    canonical_label("应收股利"): field("F013N"),
    canonical_label("买入返售金融资产"): field("F016N"),
    canonical_label("存货"): field("F015N"),
    canonical_label("合同资产"): field("F119N"),
    canonical_label("一年内到期的非流动资产"): field("F017N"),
    canonical_label("其他流动资产"): field("F018N"),
    canonical_label("流动资产合计"): field("F019N"),
    canonical_label("发放贷款和垫款"): field("F113N"),
    canonical_label("长期应收款"): field("F022N"),
    canonical_label("长期股权投资"): field("F023N"),
    canonical_label("其他权益工具投资"): field("F111N"),
    canonical_label("其他非流动金融资产"): field("F112N"),
    canonical_label("投资性房地产"): field("F024N"),
    canonical_label("固定资产"): field("F025N"),
    canonical_label("在建工程"): field("F026N"),
    canonical_label("使用权资产"): field("F121N"),
    canonical_label("无形资产"): field("F031N"),
    canonical_label("开发支出"): field("F032N"),
    canonical_label("商誉"): field("F033N"),
    canonical_label("长期待摊费用"): field("F034N"),
    canonical_label("递延所得税资产"): field("F035N"),
    canonical_label("其他非流动资产"): field("F036N"),
    canonical_label("非流动资产合计"): field("F037N"),
    canonical_label("资产总计"): field("F038N"),
    canonical_label("短期借款"): field("F039N"),
    canonical_label("交易性金融负债"): field("F090N"),
    canonical_label("衍生金融负债"): field("F091N"),
    canonical_label("应付票据"): field("F041N"),
    canonical_label("应付账款"): field("F042N"),
    canonical_label("预收款项"): field("F043N"),
    canonical_label("合同负债"): field("F115N"),
    canonical_label("应付职工薪酬"): field("F044N"),
    canonical_label("应交税费"): field("F045N"),
    canonical_label("其他应付款"): subtract_fields("F048N", "F046N", "F047N"),
    canonical_label("应付利息"): field("F046N"),
    canonical_label("应付股利"): field("F047N"),
    canonical_label("一年内到期的非流动负债"): field("F050N"),
    canonical_label("其他流动负债"): field("F051N"),
    canonical_label("流动负债合计"): field("F052N"),
    canonical_label("长期借款"): field("F053N"),
    canonical_label("应付债券"): field("F054N"),
    canonical_label("租赁负债"): field("F122N"),
    canonical_label("长期应付款"): field("F076N"),
    canonical_label("长期应付职工薪酬"): field("F055N"),
    canonical_label("预计负债"): field("F057N"),
    canonical_label("递延收益"): field("F075N"),
    canonical_label("递延所得税负债"): field("F058N"),
    canonical_label("其他非流动负债"): field("F059N"),
    canonical_label("非流动负债合计"): field("F060N"),
    canonical_label("负债合计"): field("F061N"),
    canonical_label("实收资本(或股本)"): field("F062N"),
    canonical_label("资本公积"): field("F063N"),
    canonical_label("减库存股"): field("F066N"),
    canonical_label("其他综合收益"): field("F074N"),
    canonical_label("专项储备"): field("F068N"),
    canonical_label("盈余公积"): field("F064N"),
    canonical_label("一般风险准备"): field("F069N"),
    canonical_label("未分配利润"): field("F065N"),
    canonical_label("归属于母公司所有者权益（或股东权益）合计"): field("F073N"),
    canonical_label("少数股东权益"): field("F067N"),
    canonical_label("所有者权益(或股东权益)合计"): field("F070N"),
    canonical_label("负债和所有者权益(或股东权益)总计"): field("F071N"),
}


COMPANY_INCOME_RESOLVERS: dict[str, Resolver] = {
    canonical_label("一、营业总收入"): first_available(field("F035N"), field("F006N")),
    canonical_label("营业收入"): field("F006N"),
    canonical_label("二、营业总成本"): field("F036N"),
    canonical_label("营业成本"): field("F007N"),
    canonical_label("税金及附加"): field("F008N"),
    canonical_label("销售费用"): field("F009N"),
    canonical_label("管理费用"): field("F010N"),
    canonical_label("研发费用"): field("F056N"),
    canonical_label("财务费用"): field("F012N"),
    canonical_label("加其他收益"): field("F062N"),
    canonical_label("投资收益损失以号填列"): field("F015N"),
    canonical_label("其中对联营企业和合营企业的投资收益"): field("F016N"),
    canonical_label("汇兑收益损失以填列"): field("F023N"),
    canonical_label("公允价值变动收益损失以号填列"): field("F014N"),
    canonical_label("信用减值损失损失以号填列"): field("F063N"),
    canonical_label("资产减值损失损失以号填列"): field("F064N"),
    canonical_label("资产处置收益损失以号填列"): field("F065N"),
    canonical_label("三、营业利润亏损以号填列"): field("F018N"),
    canonical_label("加营业外收入"): field("F020N"),
    canonical_label("减营业外支出"): field("F021N"),
    canonical_label("四、利润总额亏损以号填列"): field("F024N"),
    canonical_label("减所得税费用"): field("F025N"),
    canonical_label("五、净利润净亏损以号填列"): field("F027N"),
    canonical_label("1持续经营冲利润净亏损以号填列"): first_available(field("F060N"), field("F027N")),
    canonical_label("1归属于母公司股东的净利润净亏损以号填列"): field("F028N"),
    canonical_label("2少数股东损益净亏损以号填列"): field("F029N"),
    canonical_label("六、其他综合收益的税后冲额"): field("F038N"),
    canonical_label("一归属母公司所有者的其他综合收益的税后净额"): field("F066N"),
    canonical_label("二归属于少数股东的其他综合收益的税后净额"): field("F067N"),
    canonical_label("七、综合收益总额"): field("F039N"),
    canonical_label("一归属于母公司所有者的综合收益总额"): field("F040N"),
    canonical_label("二归属于少数股东的综合收益总额"): field("F041N"),
    canonical_label("一基本每股收益元股"): field("F031N"),
    canonical_label("二稀释每股收益元股"): field("F032N"),
}


COMPANY_CASH_RESOLVERS: dict[str, Resolver] = {
    canonical_label("销售商品提供劳务收到的现金"): field("F006N"),
    canonical_label("收到的税费返还"): field("F007N"),
    canonical_label("收到其他与经营活动有关的现金"): field("F008N"),
    canonical_label("经营活动现金流入小计"): field("F009N"),
    canonical_label("购买商品接受劳务支付的现金"): field("F010N"),
    canonical_label("支付给职工及为职工支付的现金"): field("F011N"),
    canonical_label("支付的各项税费"): field("F012N"),
    canonical_label("支付其他与经营活动有关的现金"): field("F013N"),
    canonical_label("经营活动现金流出小计"): field("F014N"),
    canonical_label("经营活动产生的现金流量净额"): field("F015N"),
    canonical_label("收回投资收到的现金"): field("F016N"),
    canonical_label("取得投资收益收到的现金"): field("F017N"),
    canonical_label("处置固定资产无形资产和其他长期资产收回的现金净额"): field("F018N"),
    canonical_label("处置子公司及其他营业单位收到的现金净额"): field("F019N"),
    canonical_label("收到其他与投资活动有关的现金"): field("F020N"),
    canonical_label("投资活动现金流入小计"): field("F021N"),
    canonical_label("购建固定资产无形资产和其他长期资产支付的现金"): field("F022N"),
    canonical_label("投资支付的现金"): field("F023N"),
    canonical_label("取得子公司及其他营业单位支付的现金净额"): field("F024N"),
    canonical_label("支付其他与投资活动有关的现金"): field("F025N"),
    canonical_label("投资活动现金流出小计"): field("F026N"),
    canonical_label("投资活动产生的现金流量净额"): field("F027N"),
    canonical_label("吸收投资收到的现金"): field("F028N"),
    canonical_label("其中子公司吸收少数股东投资收到的现金"): field("F089N"),
    canonical_label("取得借款收到的现金"): field("F029N"),
    canonical_label("收到其他与筹资活动有关的现金"): field("F030N"),
    canonical_label("筹资活动现金流入小计"): field("F031N"),
    canonical_label("偿还债务支付的现金"): field("F032N"),
    canonical_label("分配股利利润或偿付利息支付的现金"): field("F033N"),
    canonical_label("其中子公司支付给少数股东的股利利润"): field("F091N"),
    canonical_label("支付其他与筹资活动有关的现金"): field("F034N"),
    canonical_label("筹资活动现金流出小计"): field("F035N"),
    canonical_label("筹资活动产生的现金流量净额"): field("F036N"),
    canonical_label("四、汇率变动对现金及现金等价物的影响"): field("F037N"),
    canonical_label("现金及现金等价物净增加额"): field("F039N"),
    canonical_label("期初现金及现金等价物余额"): field("F040N"),
    canonical_label("期末现金及现金等价物余额"): field("F041N"),
}


BANK_BALANCE_RESOLVERS: dict[str, Resolver] = {
    canonical_label("现金及存放中央银行款项"): field("F006N"),
    canonical_label("衍生金融资产"): field("F035N"),
    canonical_label("买入返售款项"): field("F117N"),
    canonical_label("长期股权投资"): field("F023N"),
    canonical_label("固定资产"): field("F025N"),
    canonical_label("在建工程"): field("F026N"),
    canonical_label("递延所得税资产"): field("F087N"),
    canonical_label("资产总计"): field("F038N"),
    canonical_label("拆入资金"): field("F089N"),
    canonical_label("以公允价值计量且其变动计入当期损益的金融负债"): first_available(field("F040N"), field("F113N")),
    canonical_label("衍生金融负债"): field("F090N"),
    canonical_label("卖出回购款项"): field("F091N"),
    canonical_label("应付职工薪酬"): field("F044N"),
    canonical_label("应交税费"): field("F045N"),
    canonical_label("已发行债务证券"): field("F054N"),
    canonical_label("递延所得税负债"): field("F058N"),
    canonical_label("负债合计"): field("F061N"),
    canonical_label("股本"): field("F062N"),
    canonical_label("其他权益工具"): field("F103N"),
    canonical_label("优先股"): field("F104N"),
    canonical_label("永续债"): field("F105N"),
    canonical_label("资本公积"): field("F063N"),
    canonical_label("其他综合收益"): field("F074N"),
    canonical_label("盈余公积"): field("F064N"),
    canonical_label("一般准备"): field("F076N"),
    canonical_label("未分配利润"): field("F065N"),
    canonical_label("归属于母公司股东的权益"): field("F073N"),
    canonical_label("少数股东权益"): field("F067N"),
    canonical_label("股东权益合计"): field("F070N"),
    canonical_label("负债及股东权益总计"): field("F071N"),
}


BANK_INCOME_RESOLVERS: dict[str, Resolver] = {
    canonical_label("利息净收入"): field("F033N"),
    canonical_label("手续费及佣金净收入"): field("F042N"),
    canonical_label("投资收益"): field("F015N"),
    canonical_label("其中对联营及合营企业的投资收益"): field("F016N"),
    canonical_label("公允价值变动净收益损失"): field("F014N"),
    canonical_label("汇兑及汇率产品净损失"): field("F017N"),
    canonical_label("其他业务收入"): field("F013N"),
    canonical_label("营业收入"): field("F035N"),
    canonical_label("税金及附加"): field("F008N"),
    canonical_label("业务及管理费"): field("F057N"),
    canonical_label("资产减值损失"): field("F010N"),
    canonical_label("营业支出"): field("F036N"),
    canonical_label("营业利润"): field("F018N"),
    canonical_label("加营业外收入"): field("F020N"),
    canonical_label("减营业外支出"): field("F021N"),
    canonical_label("税前利润"): field("F024N"),
    canonical_label("减所得税费用"): field("F025N"),
    canonical_label("净利润"): field("F027N"),
    canonical_label("母公司股东"): field("F028N"),
    canonical_label("少数股东"): field("F029N"),
    canonical_label("本年净利润"): field("F027N"),
    canonical_label("其他综合收益的税后净额"): field("F038N"),
    canonical_label("本年其他综合收益小计"): field("F038N"),
    canonical_label("本年综合收益总额"): field("F039N"),
    canonical_label("基本每股收益人民币元"): field("F031N"),
    canonical_label("稀释每股收益人民币元"): field("F032N"),
}


BANK_CASH_RESOLVERS: dict[str, Resolver] = {
    canonical_label("收取的利息手续费及佣金的现金"): field("F081N"),
    canonical_label("经营活动现金流入小计"): field("F009N"),
    canonical_label("经营活动现金流出小计"): field("F014N"),
    canonical_label("经营活动产生的现金流量净额"): first_available(field("F015N"), field("F060N")),
    canonical_label("客户贷款及垫款净额"): field("F084N"),
    canonical_label("支付的利息手续费及佣金的现金"): field("F087N"),
    canonical_label("支付给职工以及为职工支付的现金支付的各项税费"): sum_fields("F011N", "F012N"),
    canonical_label("收回投资收到的现金"): field("F016N"),
    canonical_label("取得投资收益收到的现金"): field("F017N"),
    canonical_label("处置固定资产无形资产和其他长期资产不含抵债资产收回的现金"): field("F018N"),
    canonical_label("收到其他与投资活动有关的现金"): field("F020N"),
    canonical_label("投资活动现金流入小计"): field("F021N"),
    canonical_label("投资支付的现金"): field("F023N"),
    canonical_label("购建固定资产无形资产和其他长期资产支付的现金"): field("F022N"),
    canonical_label("投资活动现金流出小计"): field("F026N"),
    canonical_label("投资活动产生的现金流量净额"): field("F027N"),
    canonical_label("发行债务证券所收到的现金"): field("F076N"),
    canonical_label("筹资活动现金流入小计"): field("F031N"),
    canonical_label("筹资活动产生的现金流量净额"): field("F036N"),
    canonical_label("偿还债务证券所支付的现金"): field("F032N"),
    canonical_label("支付其他与筹资活动有关的现金"): field("F034N"),
    canonical_label("筹资活动现金流出小计"): field("F035N"),
    canonical_label("四汇率变动对现金及现金等价物的影响"): field("F037N"),
    canonical_label("现金及现金等价物净变动额"): field("F071N"),
    canonical_label("年初现金及现金等价物余额"): field("F040N"),
    canonical_label("年末现金及现金等价物余额"): field("F041N"),
}


STATEMENT_RESOLVERS: dict[tuple[str, str], dict[str, Resolver]] = {
    ("company", "balance"): COMPANY_BALANCE_RESOLVERS,
    ("company", "income"): COMPANY_INCOME_RESOLVERS,
    ("company", "cash"): COMPANY_CASH_RESOLVERS,
    ("bank", "balance"): BANK_BALANCE_RESOLVERS,
    ("bank", "income"): BANK_INCOME_RESOLVERS,
    ("bank", "cash"): BANK_CASH_RESOLVERS,
}


@dataclass(frozen=True)
class SheetLayout:
    header_row: int
    data_start_col: int
    period_count: int


@dataclass(frozen=True)
class MissingRowExplanation:
    sheet_name: str
    row_index: int
    label: str
    category: str
    detail: str


def export_template_workbook(
    company: CompanyRecord,
    balance_records: list[dict],
    income_records: list[dict],
    cash_flow_records: list[dict],
    output_dir: str | Path,
    unit_label: str = DEFAULT_UNIT_LABEL,
    template_id: str | None = None,
) -> Path:
    template = resolve_template(template_id)
    normalized_unit_label = normalize_unit_label(unit_label)
    unit_scale = UNIT_SCALE_MAP[normalized_unit_label]
    workbook = load_workbook(template.path)
    output_path = ensure_writable_dir(output_dir)
    missing_rows: list[MissingRowExplanation] = []

    if MISSING_REASON_SHEET in workbook.sheetnames:
        del workbook[MISSING_REASON_SHEET]

    balance_periods = select_annual_records(balance_records, "ENDDATE")
    income_periods = select_annual_records(income_records, "F001D")
    cash_periods = select_annual_records(cash_flow_records, "F001D")

    for sheet in workbook.worksheets:
        statement_type = detect_statement_type(sheet.title)
        if statement_type is None:
            continue

        sheet.title = build_export_sheet_title(company.secname, STATEMENT_SHEET_SUFFIX[statement_type])
        base_font = detect_sheet_base_font(sheet)
        records = {
            "balance": balance_periods,
            "income": income_periods,
            "cash": cash_periods,
        }[statement_type]
        if not records:
            continue

        layout = prepare_sheet_layout(sheet, normalized_unit_label, len(records))
        periods = records[: layout.period_count]
        write_period_headers(sheet, layout, periods)
        clear_period_cells(sheet, layout)
        missing_rows.extend(
            fill_statement_sheet(
            sheet=sheet,
            layout=layout,
            template=template,
            statement_type=statement_type,
            periods=periods,
            unit_scale=unit_scale,
            )
        )
        apply_consistent_fonts(sheet, base_font)
        apply_column_separators(sheet)
        auto_adjust_sheet_widths(sheet, layout)

    append_missing_reason_sheet(workbook, missing_rows, company.secname)

    latest_year = balance_periods[0][0][:4] if balance_periods else "latest"
    workbook_path = output_path / f"{company.secname}_{template.template_id}_{latest_year}YE.xlsx"
    workbook.save(workbook_path)
    return workbook_path


def select_annual_records(records: list[dict], date_key: str) -> list[tuple[str, dict]]:
    selected: list[tuple[str, dict]] = []
    for record in records:
        date_value = str(record.get(date_key, "")).strip()
        if not date_value.endswith("12-31"):
            continue
        if date_value != str(record.get("ENDDATE", "")).strip():
            continue
        if str(record.get("F002V", "")).strip() != "071001":
            continue
        selected.append((date_value, record))

    selected.sort(key=lambda item: item[0], reverse=True)
    return selected


def detect_statement_type(sheet_title: str) -> str | None:
    normalized = canonical_label(sheet_title)
    if "资产负债表" in normalized:
        return "balance"
    if "利润表" in normalized:
        return "income"
    if "现金流量表" in normalized:
        return "cash"
    return None


def build_export_sheet_title(company_name: str, suffix: str) -> str:
    return f"{company_name}{suffix}"[:31]


def prepare_sheet_layout(sheet, unit_label: str, record_count: int) -> SheetLayout:
    header_row = ensure_header_row(sheet, unit_label)
    data_start_col, existing_count = detect_period_columns(sheet, header_row)
    period_count = max(existing_count, record_count, DEFAULT_ANNUAL_PERIODS)
    return SheetLayout(header_row=header_row, data_start_col=data_start_col, period_count=period_count)


def ensure_header_row(sheet, unit_label: str) -> int:
    first_cell_key = canonical_label(sheet.cell(1, 1).value)
    if first_cell_key not in {canonical_label("项目"), canonical_label("报表日期")}:
        sheet.insert_rows(1)
    title_cell = sheet.cell(1, 1)
    title_cell.value = f"项目（单位：{unit_label}）"
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    title_cell.fill = HEADER_FILL
    title_cell.font = HEADER_FONT
    return 1


def detect_period_columns(sheet, header_row: int) -> tuple[int, int]:
    first_period_col: int | None = None
    existing_count = 0
    for column in range(2, sheet.max_column + 1):
        if looks_like_period_header(sheet.cell(header_row, column).value):
            first_period_col = column
            break

    if first_period_col is not None:
        column = first_period_col
        while column <= sheet.max_column and sheet.cell(header_row, column).value not in (None, ""):
            existing_count += 1
            column += 1
        return first_period_col, existing_count

    second_cell = str(sheet.cell(header_row, 2).value or "")
    if "附注" in second_cell:
        return 3, 0
    return 2, 0


def looks_like_period_header(value: object | None) -> bool:
    if isinstance(value, (date,)):
        return True
    if isinstance(value, (int, float)):
        return value > 10_000
    if isinstance(value, str):
        text = value.strip()
        return bool(DATE_PATTERN.search(text) or re.fullmatch(r"\d{4}(?:-\d{2}-\d{2})?", text))
    return False


def detect_sheet_base_font(sheet) -> Font:
    for row in range(1, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row, column)
            if cell.value not in (None, "") and cell.font is not None:
                return copy(cell.font)
    return Font(name="等线", sz=12)


def apply_consistent_fonts(sheet, base_font: Font) -> None:
    for row in range(1, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row, column)
            font = copy(cell.font)
            font.name = base_font.name
            font.sz = base_font.sz
            cell.font = font


def apply_column_separators(sheet) -> None:
    max_column = sheet.max_column
    for row in range(1, sheet.max_row + 1):
        for column in range(1, max_column + 1):
            cell = sheet.cell(row, column)
            border = copy(cell.border)
            border.left = SEPARATOR_SIDE
            if column == max_column:
                border.right = SEPARATOR_SIDE
            cell.border = border


def write_period_headers(sheet, layout: SheetLayout, periods: list[tuple[str, dict]]) -> None:
    for offset in range(layout.period_count):
        column = layout.data_start_col + offset
        cell = sheet.cell(layout.header_row, column)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        if offset < len(periods):
            period_date = date.fromisoformat(periods[offset][0])
            cell.value = period_date
            cell.number_format = "yyyy-mm-dd"
        else:
            cell.value = None


def clear_period_cells(sheet, layout: SheetLayout) -> None:
    final_col = layout.data_start_col + layout.period_count - 1
    for row in range(layout.header_row + 1, sheet.max_row + 1):
        for column in range(layout.data_start_col, final_col + 1):
            sheet.cell(row, column).value = None


def fill_statement_sheet(
    sheet,
    layout: SheetLayout,
    template: TemplateSpec,
    statement_type: str,
    periods: list[tuple[str, dict]],
    unit_scale: int,
) -> list[MissingRowExplanation]:
    resolvers = STATEMENT_RESOLVERS[(template.kind, statement_type)]
    explanations: list[MissingRowExplanation] = []

    for row in range(2, sheet.max_row + 1):
        label_cell = sheet.cell(row, 1)
        raw_label = label_cell.value
        label_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        label_key = canonical_label(raw_label)
        resolver = resolvers.get(label_key)
        apply_section_style(sheet, row, statement_type, raw_label, resolver)
        if resolver is None:
            explanation = build_missing_row_explanation(
                sheet_name=sheet.title,
                row_index=row,
                raw_label=raw_label,
                resolver=None,
                periods=periods,
            )
            if explanation is not None:
                explanations.append(explanation)
            continue

        resolved_any_value = False
        for offset, (_period, record) in enumerate(periods):
            value = resolver(record)
            if value is None:
                continue

            resolved_any_value = True
            cell = sheet.cell(row, layout.data_start_col + offset)
            if isinstance(value, (int, float)) and label_key not in NON_SCALED_LABELS:
                cell.value = value / unit_scale
                cell.number_format = "#,##0"
            else:
                cell.value = value
                if label_key in NON_SCALED_LABELS:
                    cell.number_format = "0.0000"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        if not resolved_any_value:
            explanation = build_missing_row_explanation(
                sheet_name=sheet.title,
                row_index=row,
                raw_label=raw_label,
                resolver=resolver,
                periods=periods,
            )
            if explanation is not None:
                explanations.append(explanation)

    return explanations


def apply_section_style(
    sheet,
    row_index: int,
    statement_type: str,
    raw_label: object | None,
    resolver: Resolver | None,
) -> None:
    if not is_section_label(raw_label, resolver):
        return

    fill = SECTION_FILL_BY_STATEMENT[statement_type]
    for column in range(1, sheet.max_column + 1):
        cell = sheet.cell(row_index, column)
        cell.fill = fill
        cell.font = SECTION_FONT
        if column == 1:
            cell.alignment = Alignment(horizontal="left", vertical="center")


def is_section_label(raw_label: object | None, resolver: Resolver | None) -> bool:
    text = str(raw_label or "").strip()
    if not text:
        return False
    if resolver is not None:
        return False
    if re.match(r"^[（(]\d+[)）]", text):
        return False
    if re.match(r"^\d+[、.．)]", text):
        return False
    normalized_text = text.rstrip("：:；;")
    if normalized_text in SECTION_LABEL_EXACT:
        return True
    if re.match(r"^[一二三四五六七八九十]+[、.．]", text):
        return True
    if text.endswith(("：", ":", "；", ";")) and len(normalized_text) <= 20:
        return True
    return False


def build_missing_row_explanation(
    *,
    sheet_name: str,
    row_index: int,
    raw_label: object | None,
    resolver: Resolver | None,
    periods: list[tuple[str, dict]],
) -> MissingRowExplanation | None:
    label = str(raw_label or "").strip()
    if not label:
        return None

    if resolver is None:
        if is_section_label(raw_label, resolver):
            return MissingRowExplanation(
                sheet_name=sheet_name,
                row_index=row_index,
                label=label,
                category="模板分组行",
                detail="该行用于版式分组，不对应具体取数项目。",
            )
        return MissingRowExplanation(
            sheet_name=sheet_name,
            row_index=row_index,
            label=label,
            category="待补充映射/公式",
            detail="当前程序还没有为该模板行配置取数字段或计算公式。",
        )

    source_fields = resolver_source_fields(resolver)
    if not source_fields:
        return MissingRowExplanation(
            sheet_name=sheet_name,
            row_index=row_index,
            label=label,
            category="规则未产出结果",
            detail="已有规则，但没有声明来源字段；通常意味着需要补公式或调整实现。",
        )

    present_fields = [field_name for field_name in source_fields if any(field_name in record for _, record in periods)]
    if not present_fields:
        return MissingRowExplanation(
            sheet_name=sheet_name,
            row_index=row_index,
            label=label,
            category="API 未提供字段",
            detail=f"当前接口返回中没有这些字段：{', '.join(source_fields)}。",
        )

    non_empty_fields = [
        field_name
        for field_name in source_fields
        if any(record.get(field_name) not in (None, "") for _, record in periods)
    ]
    if not non_empty_fields:
        return MissingRowExplanation(
            sheet_name=sheet_name,
            row_index=row_index,
            label=label,
            category="本期未披露数值",
            detail=f"接口字段存在，但所选年度没有披露值：{', '.join(present_fields)}。",
        )

    return MissingRowExplanation(
        sheet_name=sheet_name,
        row_index=row_index,
        label=label,
        category="已有字段但需补充计算",
        detail=(
            f"已找到来源字段 {', '.join(non_empty_fields)}，"
            f"但当前 {resolver_kind(resolver)} 规则没有产出结果，通常需要补拆分或公式。"
        ),
    )


def append_missing_reason_sheet(workbook, explanations: list[MissingRowExplanation], company_name: str) -> None:
    if not explanations:
        return

    sheet = workbook.create_sheet(build_export_sheet_title(company_name, MISSING_REASON_SHEET))
    headers = ("工作表", "行号", "项目", "原因分类", "说明")
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_index, explanation in enumerate(explanations, start=2):
        sheet.cell(row_index, 1, explanation.sheet_name)
        sheet.cell(row_index, 2, explanation.row_index)
        sheet.cell(row_index, 3, explanation.label)
        sheet.cell(row_index, 4, explanation.category)
        sheet.cell(row_index, 5, explanation.detail)

    auto_adjust_sheet_widths(sheet)


def auto_adjust_sheet_widths(sheet, layout: SheetLayout | None = None) -> None:
    max_column = sheet.max_column
    for column in range(1, max_column + 1):
        letter = get_column_letter(column)
        max_width = 0
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row, column)
            max_width = max(max_width, display_text_width(cell.value, cell.number_format))

        padding = 3 if column == 1 else 2
        width = max_width + padding
        if column == 1:
            width = min(max(width, 18), 28)
        if layout is not None and column >= layout.data_start_col:
            width = max(width, 12)
        if column == 1:
            sheet.column_dimensions[letter].width = width
        else:
            sheet.column_dimensions[letter].width = min(width, 42)

    wrap_first_column_and_adjust_row_heights(sheet)


def wrap_first_column_and_adjust_row_heights(sheet) -> None:
    first_col_width = float(sheet.column_dimensions["A"].width or 18)
    usable_width = max(first_col_width - 2, 8)

    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row, 1)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        text_width = display_text_width(cell.value, cell.number_format)
        if text_width <= 0:
            continue

        line_count = max(1, int((text_width + usable_width - 1) // usable_width))
        minimum_height = 22 if row == 1 else 20
        sheet.row_dimensions[row].height = max(minimum_height, 18 * line_count)


def display_text_width(value: object | None, number_format: str | None = None) -> int:
    if value is None:
        return 0
    if isinstance(value, date):
        text = value.isoformat()
    elif isinstance(value, (int, float)) and number_format == "#,##0":
        text = f"{value:,.0f}"
    else:
        text = str(value)
    return sum(2 if unicodedata.east_asian_width(char) in {"W", "F"} else 1 for char in text)
