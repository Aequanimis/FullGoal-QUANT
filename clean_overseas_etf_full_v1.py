from __future__ import annotations

import csv
import math
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path.home() / "Desktop" / "Fullgoal" / "课题研究"
BASE_DIR = PROJECT_ROOT / "处理后数据" / "境外策略ETF数据"
INPUT_XLSX = BASE_DIR / "原始数据" / "境外ETF全量数据 0706.xlsx"
OUTPUT_DIR = BASE_DIR / "处理结果"
OUTPUT_XLSX = OUTPUT_DIR / "境外ETF全量数据_清洗标准化_v1.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "境外ETF全量数据_清洗标准化_v1.csv"
OUTPUT_LOG = OUTPUT_DIR / "境外ETF全量数据_清洗日志_v1.txt"

EXPECTED_HEADERS = [
    "Ticker",
    "Fund Name",
    "Issuer",
    "Asset Class",
    "Category",
    "Strategy",
    "Niche",
    "Segments",
    "Selection Criteria",
    "Weighting Scheme",
    "AUM",
    "Expense Ratio",
    "Avg. Daily $ Vol.",
    "Spread Data %",
    "P/E",
    "P/B",
    "Dividend Yield",
    "1 Yr Fund Flow",
    "YTD Fund Flow",
    "1 MO Fund Flow",
    "1 YR TR",
    "3 Yr. Total Return",
    "5 Yr. Total Return",
    "ETN",
    "Leveraged",
    "Leveraged Ratio",
    "Inverse",
    "Region",
    "Geography",
    "Economic Development",
    "Avg. Daily Share Vol.",
]

RENAME_MAP = {
    "Ticker": "ticker",
    "Fund Name": "fund_name",
    "Issuer": "issuer",
    "Asset Class": "asset_class",
    "Category": "category",
    "Strategy": "strategy",
    "Niche": "niche",
    "Segments": "segments",
    "Selection Criteria": "selection_criteria",
    "Weighting Scheme": "weighting_scheme",
    "AUM": "aum_raw",
    "Expense Ratio": "expense_ratio_raw",
    "Avg. Daily $ Vol.": "avg_daily_dollar_vol_raw",
    "Spread Data %": "spread_data_pct_raw",
    "P/E": "pe_raw",
    "P/B": "pb_raw",
    "Dividend Yield": "dividend_yield_raw",
    "1 Yr Fund Flow": "one_year_fund_flow_raw",
    "YTD Fund Flow": "ytd_fund_flow_raw",
    "1 MO Fund Flow": "one_month_fund_flow_raw",
    "1 YR TR": "one_year_total_return_raw",
    "3 Yr. Total Return": "three_year_total_return_raw",
    "5 Yr. Total Return": "five_year_total_return_raw",
    "ETN": "etn_raw",
    "Leveraged": "leveraged_raw",
    "Leveraged Ratio": "leveraged_ratio",
    "Inverse": "inverse_raw",
    "Region": "region",
    "Geography": "geography",
    "Economic Development": "economic_development",
    "Avg. Daily Share Vol.": "avg_daily_share_vol_raw",
}

MASTER_FIELDS = [
    "ticker",
    "fund_name",
    "issuer",
    "asset_class",
    "category",
    "strategy",
    "niche",
    "segments",
    "selection_criteria",
    "weighting_scheme",
    "aum_raw",
    "aum_usd",
    "expense_ratio_raw",
    "expense_ratio_pct",
    "avg_daily_dollar_vol_raw",
    "avg_daily_dollar_vol_usd",
    "avg_daily_share_vol_raw",
    "avg_daily_share_vol",
    "spread_data_pct_raw",
    "spread_data_pct",
    "pe_raw",
    "pe",
    "pb_raw",
    "pb",
    "dividend_yield_raw",
    "dividend_yield_pct",
    "one_year_fund_flow_raw",
    "one_year_fund_flow_usd",
    "ytd_fund_flow_raw",
    "ytd_fund_flow_usd",
    "one_month_fund_flow_raw",
    "one_month_fund_flow_usd",
    "one_year_total_return_raw",
    "one_year_total_return_pct",
    "three_year_total_return_raw",
    "three_year_total_return_pct",
    "five_year_total_return_raw",
    "five_year_total_return_pct",
    "etn_raw",
    "etn",
    "leveraged_raw",
    "leveraged",
    "leveraged_ratio",
    "inverse_raw",
    "inverse",
    "region",
    "geography",
    "economic_development",
    "preliminary_pool_type",
    "preliminary_strategy_level_1",
    "preliminary_strategy_level_2",
    "classification_reason",
    "is_excluded_by_features",
    "is_strategy_candidate",
    "is_extended_strategy_candidate",
    "ticker_format_warning",
    "missing_key_field_warning",
    "numeric_parse_warning",
    "classification_warning",
    "source_file",
    "extract_time",
]

AMOUNT_RAW_FIELDS = [
    "aum_raw",
    "avg_daily_dollar_vol_raw",
    "avg_daily_share_vol_raw",
    "one_year_fund_flow_raw",
    "ytd_fund_flow_raw",
    "one_month_fund_flow_raw",
]
AMOUNT_CLEAN_MAP = {
    "aum_raw": "aum_usd",
    "avg_daily_dollar_vol_raw": "avg_daily_dollar_vol_usd",
    "avg_daily_share_vol_raw": "avg_daily_share_vol",
    "one_year_fund_flow_raw": "one_year_fund_flow_usd",
    "ytd_fund_flow_raw": "ytd_fund_flow_usd",
    "one_month_fund_flow_raw": "one_month_fund_flow_usd",
}
PCT_CLEAN_MAP = {
    "expense_ratio_raw": "expense_ratio_pct",
    "spread_data_pct_raw": "spread_data_pct",
    "dividend_yield_raw": "dividend_yield_pct",
    "one_year_total_return_raw": "one_year_total_return_pct",
    "three_year_total_return_raw": "three_year_total_return_pct",
    "five_year_total_return_raw": "five_year_total_return_pct",
}
VALUATION_CLEAN_MAP = {"pe_raw": "pe", "pb_raw": "pb"}
BOOL_CLEAN_MAP = {"etn_raw": "etn", "leveraged_raw": "leveraged", "inverse_raw": "inverse"}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def is_missing_like(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    text = normalize_text(value)
    return text == "" or text in {"-", "--", "N/A", "nan", "NaN"} or text.lower() == "undefined%"


def format_raw(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float):
        return value
    return normalize_text(value)


def parse_amount(value: Any) -> Tuple[Optional[float], Optional[str]]:
    if is_missing_like(value):
        return None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), None
    text = normalize_text(value).replace("−", "-").replace("—", "-").replace("–", "-")
    if text.lower() == "undefined%":
        return None, None
    m = re.fullmatch(r"(-)?\$?\s*([\d,]+(?:\.\d+)?)\s*([KMBkmb]?)", text)
    if not m:
        return None, f"金额字段无法解析: {text}"
    sign = -1 if m.group(1) else 1
    num = float(m.group(2).replace(",", ""))
    multiplier = {"": 1, "K": 1_000, "M": 1_000_000, "B": 1_000_000_000}[m.group(3).upper()]
    return sign * num * multiplier, None


def parse_percent_point(value: Any) -> Tuple[Optional[float], Optional[str]]:
    if is_missing_like(value):
        return None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) * 100, None
    text = normalize_text(value).replace("−", "-").replace("—", "-").replace("–", "-")
    if text.lower() == "undefined%":
        return None, None
    if text.endswith("%"):
        core = text[:-1].replace(",", "").strip()
        try:
            return float(core), None
        except ValueError:
            return None, f"百分比字段无法解析: {text}"
    try:
        # Text without % is treated the same as Excel numeric decimal.
        return float(text.replace(",", "")) * 100, None
    except ValueError:
        return None, f"百分比字段无法解析: {text}"


def parse_number(value: Any) -> Tuple[Optional[float], Optional[str]]:
    if is_missing_like(value):
        return None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), None
    text = normalize_text(value)
    try:
        return float(text.replace(",", "")), None
    except ValueError:
        return None, f"估值字段无法解析: {text}"


def parse_bool(value: Any) -> Tuple[Optional[bool], Optional[str]]:
    if is_missing_like(value):
        return None, None
    text = normalize_text(value).lower()
    if text == "yes":
        return True, None
    if text == "no":
        return False, None
    return None, f"布尔字段无法解析: {value}"


def find_keyword_hits(row: Dict[str, Any], fields: Iterable[str], keywords: List[str]) -> List[Tuple[str, str]]:
    hits: List[Tuple[str, str]] = []
    for field in fields:
        value = str(row.get(field, "") or "")
        lower = value.lower()
        for kw in keywords:
            if kw.lower() in lower:
                hits.append((field, kw))
    return hits


def classify_row(row: Dict[str, Any]) -> None:
    row["preliminary_strategy_level_1"] = ""
    row["preliminary_strategy_level_2"] = ""
    row["classification_reason"] = ""
    row["is_excluded_by_features"] = False
    row["is_strategy_candidate"] = False
    row["is_extended_strategy_candidate"] = False

    fund_name = str(row.get("fund_name", "") or "")
    segments = str(row.get("segments", "") or "")
    leveraged_ratio = normalize_text(row.get("leveraged_ratio"))
    exclusion_hits: List[str] = []
    if row.get("leveraged") is True:
        exclusion_hits.append("leveraged=True")
    if row.get("inverse") is True:
        exclusion_hits.append("inverse=True")
    if row.get("etn") is True:
        exclusion_hits.append("etn=True")
    if leveraged_ratio and leveraged_ratio != "-":
        exclusion_hits.append(f"leveraged_ratio={leveraged_ratio}")
    if re.search(r"\b(Leveraged|Inverse)\b", segments, flags=re.I):
        exclusion_hits.append("segments contains Leveraged/Inverse")
    if re.search(r"\b(2X|3X|Ultra|Bear|Bull 2X|Daily Target|Daily ETF)\b", fund_name, flags=re.I):
        exclusion_hits.append("fund_name contains leveraged/inverse keyword")
    if exclusion_hits:
        row["preliminary_pool_type"] = "排除_杠杆反向ETN"
        row["is_excluded_by_features"] = True
        row["classification_reason"] = "; ".join(exclusion_hits)
        return

    extended_keywords = [
        "Structured Outcome",
        "Buffer",
        "Defined Protection",
        "Defined Outcome",
        "Autocallable",
        "Barrier",
        "Income Strategy",
        "Yield Premium",
        "Premium Income",
        "Covered Call",
        "BuyWrite",
        "Collared",
        "Laddered Buffer",
        "Downside Hedge",
        "Target Outcome",
        "Capital Appreciation",
        "Target Risk",
    ]
    ext_hits = find_keyword_hits(row, ["category", "strategy", "niche", "segments", "fund_name"], extended_keywords)
    if ext_hits:
        combined = " ".join(kw for _, kw in ext_hits).lower()
        if any(k in combined for k in ["buffer", "defined protection", "defined outcome", "structured outcome"]):
            lvl1 = "Buffer/Defined Outcome"
        elif any(k in combined for k in ["autocallable", "barrier"]):
            lvl1 = "Autocallable/Barrier Income"
        elif any(k in combined for k in ["income strategy", "yield premium", "premium income", "covered call", "buywrite"]):
            lvl1 = "Option Income/Premium Income"
        elif any(k in combined for k in ["downside hedge", "protection"]):
            lvl1 = "Downside Hedge/Protection"
        elif any(k in combined for k in ["capital appreciation", "target risk", "target outcome"]):
            lvl1 = "Asset Allocation/Target Outcome"
        else:
            lvl1 = "Other Extended Strategy"
        row["preliminary_pool_type"] = "扩展策略ETF"
        row["is_strategy_candidate"] = True
        row["is_extended_strategy_candidate"] = True
        row["preliminary_strategy_level_1"] = lvl1
        row["preliminary_strategy_level_2"] = lvl1
        row["classification_reason"] = "命中扩展策略关键词：" + "; ".join(f"{f}:{k}" for f, k in ext_hits[:10])
        return

    core_priority = [
        ("Dividend", ["High Dividend Yield", "Dividend"]),
        ("Free Cash Flow", ["Free Cash Flow", "FCF", "Cash Cow"]),
        ("Low Volatility", ["Low Vol", "Min Vol", "Minimum Volatility", "Volatility"]),
        ("Quality", ["Quality"]),
        ("Value", ["Value"]),
        ("Growth", ["Growth"]),
        ("Momentum", ["Momentum"]),
        ("Factor / Multi-Factor", ["Multi-Factor", "Active Factor", "Alpha Capture", "Factor"]),
        ("Fundamental Weighting", ["Fundamental"]),
        ("Shareholder Yield", ["Shareholder Yield", "Buyback"]),
        ("Moat", ["Moat"]),
        ("Equal Weight", ["Equal"]),
        ("Yield Enhanced", ["Yield Enhanced"]),
        ("Sector Rotation", ["Sector Rotation"]),
        ("Revenue Weighting", ["Revenue"]),
        ("Earnings Weighting", ["Earnings"]),
        ("Liquidity Screen", ["Liquidity"]),
        ("Technical Strategy", ["Technical"]),
        ("Bottom-up Strategy", ["Bottom-up"]),
        ("Alternative Weighting", ["Arbitrary Weighting"]),
    ]
    core_fields = [
        "fund_name",
        "category",
        "strategy",
        "niche",
        "segments",
        "selection_criteria",
        "weighting_scheme",
    ]
    selected_lvl1 = ""
    core_hits: List[Tuple[str, str]] = []
    style_exposure_note = ""
    category = normalize_text(row.get("category"))
    niche = normalize_text(row.get("niche"))
    strategy = normalize_text(row.get("strategy"))
    if category == "Size and Style" and niche in {"Growth", "Value"}:
        selected_lvl1 = niche
        core_hits = [("category/niche", f"Size and Style + {niche}")]
        style_exposure_note = "风格暴露类，需后续与普通市值加权风格ETF区分"
    else:
        for lvl1, keywords in core_priority:
            hits = find_keyword_hits(row, core_fields, keywords)
            if hits:
                selected_lvl1 = lvl1
                core_hits.extend(hits)
                break
        for _, keywords in core_priority:
            for hit in find_keyword_hits(row, core_fields, keywords):
                if hit not in core_hits:
                    core_hits.append(hit)
    if selected_lvl1:
        row["preliminary_pool_type"] = "核心策略ETF_候选"
        row["is_strategy_candidate"] = True
        row["is_extended_strategy_candidate"] = False
        row["preliminary_strategy_level_1"] = selected_lvl1
        row["preliminary_strategy_level_2"] = selected_lvl1
        reason = "命中核心策略关键词：" + "; ".join(f"{f}:{k}" for f, k in core_hits[:12])
        if style_exposure_note:
            reason += f"；{style_exposure_note}"
        row["classification_reason"] = reason
        return

    selection = normalize_text(row.get("selection_criteria"))
    weighting = normalize_text(row.get("weighting_scheme"))
    if strategy == "Active" or selection == "Proprietary" or weighting == "Proprietary":
        row["preliminary_pool_type"] = "观察池_主动或专有策略待核验"
        row["classification_reason"] = "Active 或 Proprietary 口径但未命中核心/扩展策略关键词"
        return

    asset_class = normalize_text(row.get("asset_class"))
    if (
        (category == "Sector" and strategy == "Vanilla")
        or asset_class == "Commodities"
        or asset_class == "Currency"
        or (
            asset_class == "Fixed Income"
            and any(key in category for key in ["Broad Market", "Corporate", "Bank Loans", "Government"])
        )
        or (
            selection == "Market Cap"
            and weighting == "Market Cap"
            and strategy == "Vanilla"
            and niche == "Broad-based"
        )
    ):
        row["preliminary_pool_type"] = "排除_普通宽基主题商品债券"
        row["classification_reason"] = "未命中策略关键词，符合普通宽基/主题/商品/债券排除规则"
        return

    row["preliminary_pool_type"] = "待人工复核"
    row["classification_reason"] = "未命中排除、核心策略、扩展策略或主动/专有观察规则"


def build_warnings(row: Dict[str, Any], parse_warnings: List[str]) -> None:
    ticker = normalize_text(row.get("ticker"))
    row["ticker_format_warning"] = "" if re.fullmatch(r"[A-Z0-9]{1,8}", ticker or "") else "特殊ticker格式"
    missing_keys = [field for field in ["ticker", "fund_name", "asset_class", "category"] if not normalize_text(row.get(field))]
    row["missing_key_field_warning"] = "缺失关键字段: " + ", ".join(missing_keys) if missing_keys else ""
    row["numeric_parse_warning"] = "；".join(parse_warnings)
    missing_class = [field for field in ["strategy", "segments", "issuer"] if not normalize_text(row.get(field))]
    row["classification_warning"] = "分类参考字段缺失: " + ", ".join(missing_class) if missing_class else ""


def read_input_rows() -> Tuple[List[Dict[str, Any]], int]:
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"输入文件不存在: {INPUT_XLSX}")
    wb = load_workbook(INPUT_XLSX, read_only=True, data_only=True)
    if wb.sheetnames != ["Sheet1"]:
        raise ValueError(f"输入文件 sheet 异常，预期仅 Sheet1，实际: {wb.sheetnames}")
    ws = wb["Sheet1"]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    if headers != EXPECTED_HEADERS:
        missing = [h for h in EXPECTED_HEADERS if h not in headers]
        extra = [h for h in headers if h not in EXPECTED_HEADERS]
        wb.close()
        raise ValueError(f"输入字段发生变化，停止处理。missing={missing}; extra={extra}; actual={headers}")
    rows: List[Dict[str, Any]] = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        raw = dict(zip(headers, values))
        row = {RENAME_MAP[k]: format_raw(v) for k, v in raw.items()}
        rows.append(row)
    wb.close()
    return rows, ws.max_row - 1


def process_rows(raw_rows: List[Dict[str, Any]], extract_time: str) -> List[Dict[str, Any]]:
    processed: List[Dict[str, Any]] = []
    for raw in raw_rows:
        row = {field: raw.get(field, "") for field in RENAME_MAP.values()}
        parse_warnings: List[str] = []
        for raw_field, clean_field in AMOUNT_CLEAN_MAP.items():
            val, warning = parse_amount(row.get(raw_field))
            row[clean_field] = val
            if warning:
                parse_warnings.append(f"{raw_field}: {warning}")
        for raw_field, clean_field in PCT_CLEAN_MAP.items():
            val, warning = parse_percent_point(row.get(raw_field))
            row[clean_field] = val
            if warning:
                parse_warnings.append(f"{raw_field}: {warning}")
        for raw_field, clean_field in VALUATION_CLEAN_MAP.items():
            val, warning = parse_number(row.get(raw_field))
            row[clean_field] = val
            if warning:
                parse_warnings.append(f"{raw_field}: {warning}")
        for raw_field, clean_field in BOOL_CLEAN_MAP.items():
            val, warning = parse_bool(row.get(raw_field))
            row[clean_field] = val
            if warning:
                parse_warnings.append(f"{raw_field}: {warning}")
        classify_row(row)
        build_warnings(row, parse_warnings)
        row["source_file"] = str(INPUT_XLSX)
        row["extract_time"] = extract_time
        processed.append({field: row.get(field, "") for field in MASTER_FIELDS})
    return processed


def sort_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return sorted(rows, key=lambda r: (str(r.get("preliminary_pool_type", "")), str(r.get("ticker", ""))))


def excel_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def write_table_sheet(ws, rows: List[Dict[str, Any]], headers: List[str]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([excel_value(row.get(h)) for h in headers])
    style_sheet(ws)
    ws.auto_filter.ref = ws.dimensions


def style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    numeric_money = {
        "aum_usd",
        "avg_daily_dollar_vol_usd",
        "avg_daily_share_vol",
        "one_year_fund_flow_usd",
        "ytd_fund_flow_usd",
        "one_month_fund_flow_usd",
    }
    pct_fields = {
        "expense_ratio_pct",
        "spread_data_pct",
        "dividend_yield_pct",
        "one_year_total_return_pct",
        "three_year_total_return_pct",
        "five_year_total_return_pct",
    }
    valuation_fields = {"pe", "pb"}
    headers = [c.value for c in ws[1]]
    for col_idx, field in enumerate(headers, 1):
        letter = get_column_letter(col_idx)
        max_len = len(str(field or ""))
        for cell in ws[letter][1:]:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)), 80))
            if field in numeric_money:
                cell.number_format = "#,##0"
            elif field in pct_fields or field in valuation_fields:
                cell.number_format = "0.00"
            elif field == "ticker":
                cell.number_format = "@"
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 45)


def counter_table(title: str, counter: Counter) -> List[List[Any]]:
    rows = [[title, "数量"]]
    for key, val in counter.most_common():
        rows.append([key if key not in (None, "") else "(空)", val])
    return rows


def build_quality_report(rows: List[Dict[str, Any]], raw_count: int) -> List[List[Any]]:
    tickers = [r.get("ticker") for r in rows if r.get("ticker")]
    dup_count = len(tickers) - len(set(tickers))
    report: List[List[Any]] = [
        ["指标", "数值"],
        ["原始行数", raw_count],
        ["清洗后行数", len(rows)],
        ["唯一 ticker 数", len(set(tickers))],
        ["重复 ticker 数", dup_count],
        ["字段数", len(MASTER_FIELDS)],
        ["缺失 issuer 数", sum(not normalize_text(r.get("issuer")) for r in rows)],
        ["缺失 AUM 数", sum(is_missing_like(r.get("aum_raw")) for r in rows)],
        ["undefined expense ratio 数", sum(str(r.get("expense_ratio_raw")).lower() == "undefined%" for r in rows)],
        ["缺失 strategy 数", sum(not normalize_text(r.get("strategy")) for r in rows)],
        ["缺失 segments 数", sum(not normalize_text(r.get("segments")) for r in rows)],
        ["P/E 缺失数", sum(is_missing_like(r.get("pe_raw")) for r in rows)],
        ["P/B 缺失数", sum(is_missing_like(r.get("pb_raw")) for r in rows)],
        ["Dividend Yield 缺失数", sum(is_missing_like(r.get("dividend_yield_raw")) for r in rows)],
        ["ETN=Yes 数", sum(r.get("etn") is True for r in rows)],
        ["Leveraged=Yes 数", sum(r.get("leveraged") is True for r in rows)],
        ["Inverse=Yes 数", sum(r.get("inverse") is True for r in rows)],
        ["ticker_format_warning 非空数量", sum(bool(r.get("ticker_format_warning")) for r in rows)],
        ["missing_key_field_warning 非空数量", sum(bool(r.get("missing_key_field_warning")) for r in rows)],
        ["numeric_parse_warning 非空数量", sum(bool(r.get("numeric_parse_warning")) for r in rows)],
        ["classification_warning 非空数量", sum(bool(r.get("classification_warning")) for r in rows)],
    ]
    for title, field in [
        ("各 asset_class 数量分布", "asset_class"),
        ("各 category 数量分布", "category"),
        ("各 strategy 数量分布", "strategy"),
        ("各 preliminary_pool_type 数量分布", "preliminary_pool_type"),
        ("各 preliminary_strategy_level_1 数量分布", "preliminary_strategy_level_1"),
    ]:
        report.append([])
        report.extend(counter_table(title, Counter(r.get(field) for r in rows)))
    return report


def write_matrix_sheet(ws, matrix: List[List[Any]]) -> None:
    for row in matrix:
        ws.append(row)
    style_sheet(ws)


def field_dictionary_rows() -> List[List[Any]]:
    desc = {
        "ticker": ("ETF 代码", "Ticker", "保留文本", "无"),
        "fund_name": ("基金名称", "Fund Name", "去除首尾空格", "无"),
        "issuer": ("发行人/管理人", "Issuer", "去除首尾空格", "无"),
        "aum_usd": ("资产管理规模", "AUM", "K/M/B 转美元数值", "美元"),
        "expense_ratio_pct": ("费率", "Expense Ratio", "转为百分点", "百分点"),
        "avg_daily_dollar_vol_usd": ("日均成交额", "Avg. Daily $ Vol.", "K/M/B 转美元数值", "美元"),
        "avg_daily_share_vol": ("日均成交份额", "Avg. Daily Share Vol.", "K/M/B 转数值", "份"),
        "spread_data_pct": ("买卖价差", "Spread Data %", "转为百分点", "百分点"),
        "pe": ("市盈率", "P/E", "转 float", "倍"),
        "pb": ("市净率", "P/B", "转 float", "倍"),
        "dividend_yield_pct": ("股息率", "Dividend Yield", "转为百分点", "百分点"),
        "preliminary_pool_type": ("初步产品池分类", "衍生字段", "按关键词和排除规则生成", "无"),
    }
    rows = [["字段", "中文含义", "原始字段来源", "清洗规则", "单位"]]
    for field in MASTER_FIELDS:
        if field in desc:
            rows.append([field, *desc[field]])
        elif field.endswith("_raw"):
            rows.append([field, "原始字段", field.replace("_raw", ""), "保留原始值", "原始单位"])
        elif field.endswith("_warning"):
            rows.append([field, "质量检查提示", "衍生字段", "异常时写入提示", "无"])
        else:
            rows.append([field, "清洗/分类字段", "原始或衍生字段", "见清洗脚本规则", "视字段而定"])
    return rows


def classification_rules_rows() -> List[List[Any]]:
    return [
        ["规则模块", "条件/关键词", "输出分类", "注意事项"],
        ["排除_杠杆反向ETN", "Leveraged=True; Inverse=True; ETN=True; leveraged_ratio 非空; segments 含 Leveraged/Inverse; fund_name 含 2X/3X/Ultra/Bear/Daily ETF", "排除_杠杆反向ETN", "优先级最高"],
        ["扩展策略ETF", "Structured Outcome, Buffer, Defined Protection, Autocallable, Covered Call, Premium Income, Target Outcome 等", "扩展策略ETF", "期权收益、buffer、目标结果类"],
        ["核心策略ETF_候选", "Dividend, Value, Growth, Quality, Low Vol, Momentum, Factor, Fundamental, Equal, FCF, Moat 等", "核心策略ETF_候选", "只做初筛，后续仍需复核"],
        ["风格暴露特别规则", "category=Size and Style 且 niche=Growth/Value", "核心策略ETF_候选", "标注风格暴露类需后续复核"],
        ["观察池", "Active 或 Proprietary，但未命中核心/扩展关键词", "观察池_主动或专有策略待核验", "不直接纳入核心统计"],
        ["普通产品排除", "普通行业、商品、货币、普通债券、普通宽基市值加权", "排除_普通宽基主题商品债券", "没有策略关键词时排除"],
        ["兜底", "以上均未命中", "待人工复核", "人工判断"],
    ]


def explanation_rows(raw_count: int, clean_count: int, extract_time: str) -> List[List[Any]]:
    return [
        ["境外ETF全量数据清洗标准化 v1"],
        ["数据来源", "ETF.com 手动导出 Excel"],
        ["输入文件", str(INPUT_XLSX)],
        ["输出文件", str(OUTPUT_XLSX)],
        ["处理时间", extract_time],
        ["原始行数", raw_count],
        ["清洗后行数", clean_count],
        ["主要处理逻辑", "字段 snake_case 标准化；金额/百分比/估值/布尔字段转换；质量 warning；初步策略分类；分池输出"],
        ["注意事项", "本版本为初步规则分类，核心策略候选、扩展策略和观察池均需结合后续研究口径复核。"],
    ]


def write_outputs(rows: List[Dict[str, Any]], raw_count: int, extract_time: str) -> None:
    with OUTPUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=MASTER_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "00_说明"
    write_matrix_sheet(ws, explanation_rows(raw_count, len(rows), extract_time))

    sheets = [
        ("01_master_clean", rows),
        ("02_core_strategy_candidate", [r for r in rows if r.get("preliminary_pool_type") == "核心策略ETF_候选"]),
        ("03_extended_strategy_candidate", [r for r in rows if r.get("preliminary_pool_type") == "扩展策略ETF"]),
        ("04_observation_pool", [r for r in rows if "观察池" in str(r.get("preliminary_pool_type")) or r.get("preliminary_pool_type") == "待人工复核"]),
        ("05_excluded_pool", [r for r in rows if str(r.get("preliminary_pool_type", "")).startswith("排除")]),
    ]
    for name, data in sheets:
        ws = wb.create_sheet(name)
        write_table_sheet(ws, data, MASTER_FIELDS)

    ws = wb.create_sheet("06_quality_report")
    write_matrix_sheet(ws, build_quality_report(rows, raw_count))

    ws = wb.create_sheet("07_field_dictionary")
    write_matrix_sheet(ws, field_dictionary_rows())

    ws = wb.create_sheet("08_classification_rules")
    write_matrix_sheet(ws, classification_rules_rows())

    wb.save(OUTPUT_XLSX)


def sample_lines(rows: List[Dict[str, Any]]) -> List[str]:
    sample_tickers = [
        "ACWV",
        "ABFL",
        "AAPR",
        "ABUF",
        "AALG",
        "AAPU",
        "AGQ",
        "AGIQ",
        "AGIX",
        "VUG",
        "SCHD",
        "COWZ",
        "QUAL",
        "MTUM",
        "USMV",
        "MOAT",
        "JEPI",
    ]
    by_ticker = {r.get("ticker"): r for r in rows}
    lines: List[str] = []
    for ticker in sample_tickers:
        row = by_ticker.get(ticker)
        if not row:
            lines.append(f"- {ticker}: 不在本次数据中")
        else:
            lines.append(
                f"- {ticker}: {row.get('fund_name')} | {row.get('preliminary_pool_type')} | "
                f"{row.get('preliminary_strategy_level_1')} | {row.get('classification_reason')}"
            )
    return lines


def write_log(rows: List[Dict[str, Any]], raw_count: int, extract_time: str) -> str:
    tickers = [r.get("ticker") for r in rows if r.get("ticker")]
    lines = [
        "境外ETF全量数据 清洗日志 v1",
        f"处理时间: {extract_time}",
        f"输入文件: {INPUT_XLSX}",
        f"输出 Excel: {OUTPUT_XLSX}",
        f"输出 CSV: {OUTPUT_CSV}",
        f"清空的目标目录: {OUTPUT_DIR}",
        f"原始行数: {raw_count}",
        f"清洗后行数: {len(rows)}",
        f"唯一 ticker 数: {len(set(tickers))}",
        f"重复 ticker 数: {len(tickers) - len(set(tickers))}",
        "各 preliminary_pool_type 数量:",
    ]
    for k, v in Counter(r.get("preliminary_pool_type") for r in rows).most_common():
        lines.append(f"  - {k}: {v}")
    lines.extend(
        [
            "主要 warning 统计:",
            f"  - ticker_format_warning 非空: {sum(bool(r.get('ticker_format_warning')) for r in rows)}",
            f"  - missing_key_field_warning 非空: {sum(bool(r.get('missing_key_field_warning')) for r in rows)}",
            f"  - numeric_parse_warning 非空: {sum(bool(r.get('numeric_parse_warning')) for r in rows)}",
            f"  - classification_warning 非空: {sum(bool(r.get('classification_warning')) for r in rows)}",
            "重点样本分类结果:",
            *sample_lines(rows),
        ]
    )
    text = "\n".join(lines)
    OUTPUT_LOG.write_text(text, encoding="utf-8-sig")
    return text


def main() -> None:
    print(f"输入文件: {INPUT_XLSX}")
    print(f"输出目录: {OUTPUT_DIR}")
    raw_rows, raw_count = read_input_rows()
    extract_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = process_rows(raw_rows, extract_time)
    write_outputs(rows, raw_count, extract_time)
    log_text = write_log(rows, raw_count, extract_time)
    tickers = [r.get("ticker") for r in rows if r.get("ticker")]
    print("\n验收摘要")
    print(f"原始行数: {raw_count}")
    print(f"清洗后行数: {len(rows)}")
    print(f"唯一 ticker 数: {len(set(tickers))}")
    print(f"重复 ticker 数: {len(tickers) - len(set(tickers))}")
    print("preliminary_pool_type 分布:")
    for k, v in Counter(r.get("preliminary_pool_type") for r in rows).most_common():
        print(f"  {k}: {v}")
    print(f"ticker_format_warning 非空: {sum(bool(r.get('ticker_format_warning')) for r in rows)}")
    print(f"missing_key_field_warning 非空: {sum(bool(r.get('missing_key_field_warning')) for r in rows)}")
    print(f"numeric_parse_warning 非空: {sum(bool(r.get('numeric_parse_warning')) for r in rows)}")
    print(f"classification_warning 非空: {sum(bool(r.get('classification_warning')) for r in rows)}")
    print("重点样本:")
    for line in sample_lines(rows):
        print(line)
    print(f"输出 Excel: {OUTPUT_XLSX}")
    print(f"输出 CSV: {OUTPUT_CSV}")
    print(f"输出日志: {OUTPUT_LOG}")


if __name__ == "__main__":
    main()
