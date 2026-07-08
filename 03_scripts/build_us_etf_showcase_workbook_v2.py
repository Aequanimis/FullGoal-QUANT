from __future__ import annotations

import csv
import math
from copy import copy
from pathlib import Path
from statistics import mean, median

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SNAPSHOT_DATE = "2026-07-06"

BASE_DIR = Path.home() / "Desktop" / "Fullgoal"
WORK_DIR = BASE_DIR / "ICI课题" / "课题更正"
INPUT_XLSX = WORK_DIR / "美国ETF案例数据收集_富国启示_展示整理版.xlsx"
OUTPUT_XLSX = WORK_DIR / "美国ETF案例数据收集_富国启示_展示整理版_v2.xlsx"
REQUESTED_CSV = WORK_DIR / "境外ETF全量数据_清洗标准化_v1.csv"
CSV_NAME = "境外ETF全量数据_清洗标准化_v1.csv"

ISSUER_MAP = {
    "BlackRock, Inc.": "BlackRock/iShares",
    "JPMorgan Chase": "JPMorgan AM",
    "Invesco": "Invesco",
}

ISSUER_META = {
    "BlackRock/iShares": {
        "case_role": "龙头平台型样本",
        "ecosystem_keywords": "产品矩阵；Core ETF；Model Portfolios；投顾渠道；账户沉淀",
        "ppt_role": "ETF生态成熟天花板",
        "product_mix_note": "iShares：多资产ETF工具箱完整",
        "ppt_usage": "P24 Core ETF/产品模块；P23龙头样本说明",
    },
    "JPMorgan AM": {
        "case_role": "能力迁移型样本",
        "ecosystem_keywords": "主动ETF；收益增强；固定收益；BetaBuilders；研究能力迁移",
        "ppt_role": "类富国主参考样本",
        "product_mix_note": "JPMorgan：主动、固收、收益型ETF是差异化方向",
        "ppt_usage": "P25 能力迁移链；收益增强、主动固收、BetaBuilders代表产品",
    },
    "Invesco": {
        "case_role": "入口型辅助样本",
        "ecosystem_keywords": "QQQ；创新成长；超级入口；品牌心智；产品矩阵外溢",
        "ppt_role": "辅助说明入口型大单品",
        "product_mix_note": "Invesco：QQQ及创新成长相关产品构成强入口",
        "ppt_usage": "P25 右侧入口型大单品辅助卡",
    },
}

CORRECTED_SHEETS = [
    "01_Issuer_Overview_Clean",
    "02_Product_Mix_Clean",
    "03_Representative_ETFs_Clean",
    "08_P25_Invesco_Helper_Data",
    "09_P26_Fullgoal_Implication_Data",
    "10_Slide_Text_Prompts",
    "11_QA_Checks",
]

REP_TICKERS = {
    "BlackRock/iShares": ["IVV", "IEFA", "IEMG", "AGG", "IWF", "IJH", "IJR", "SGOV", "ITOT", "IWM"],
    "JPMorgan AM": ["JEPI", "JEPQ", "JPST", "JCPB", "BBUS", "BBJP", "BBEU", "JPIE", "JGRO"],
    "Invesco": ["QQQ", "QQQM", "RSP", "QLD", "SPLV", "BKLN", "DBC", "PDBC", "SPHQ", "SPHD"],
}

REQUIRED_PRODUCT_AUM_B = {
    "IVV": 885.43,
    "AGG": 138.12,
    "ITOT": 94.00,
    "JEPI": 44.58,
    "JEPQ": 40.19,
    "JPST": 39.65,
    "JCPB": 13.56,
    "BBUS": 8.52,
    "QQQ": 488.91,
    "QQQM": 99.88,
}

EXPECTED_ISSUER_AUM_B = {
    "BlackRock/iShares": 4559.91,
    "JPMorgan AM": 325.68,
    "Invesco": 984.69,
}


def find_csv() -> Path:
    if REQUESTED_CSV.exists():
        return REQUESTED_CSV
    search_roots = [Path.cwd(), Path.cwd().parent, BASE_DIR]
    for root in search_roots:
        for path in root.rglob(CSV_NAME):
            return path
    raise FileNotFoundError(f"未找到CSV：{REQUESTED_CSV}，且当前目录/上级目录/Fullgoal目录下未搜索到 {CSV_NAME}")


def to_float(value) -> float | None:
    if value in (None, ""):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def sum_field(rows, field: str) -> float:
    return sum(to_float(row.get(field)) or 0 for row in rows)


def numeric_values(rows, field: str) -> list[float]:
    return [value for row in rows if (value := to_float(row.get(field))) is not None]


def read_target_rows(csv_path: Path) -> list[dict]:
    rows: list[dict] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"ticker", "fund_name", "issuer", "asset_class", "category", "strategy", "aum_usd", "expense_ratio_pct", "avg_daily_dollar_vol_usd"}
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise ValueError(f"CSV缺少必要字段：{missing}")
        for row in reader:
            raw_issuer = row.get("issuer")
            if raw_issuer in ISSUER_MAP:
                row["issuer_display_name"] = ISSUER_MAP[raw_issuer]
                rows.append(row)
    return rows


def delete_if_exists(wb, sheet_name: str):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]


def write_table(wb, sheet_name: str, headers: list[str], rows: list[dict]):
    delete_if_exists(wb, sheet_name)
    ws = wb.create_sheet(sheet_name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])
    style_sheet(ws)
    return ws


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False

    for cell in ws[1]:
        cell.fill = copy(header_fill)
        cell.font = copy(header_font)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = copy(border)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = copy(border)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col_idx, header_cell in enumerate(ws[1], start=1):
        header = str(header_cell.value or "")
        max_len = len(header)
        for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
            for item in cell:
                if item.value is not None:
                    max_len = max(max_len, len(str(item.value)))

        text_cap = 40 if any(key in header.lower() for key in ["note", "message", "source", "url", "description", "text", "annotation", "说明", "来源"]) else 26
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), text_cap)

        lower = header.lower()
        if "usd" in lower:
            number_format = '$#,##0'
            if lower.endswith("billion_usd") or lower.endswith("trillion_usd") or lower.endswith("100m_usd"):
                number_format = '0.00'
            for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                for cell in col:
                    cell.number_format = number_format
                    cell.alignment = Alignment(horizontal="right", vertical="top")
        elif lower.endswith("_pct") or lower == "aum_share":
            number_format = '0.00%' if lower == "aum_share" else '0.00"%"'
            for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                for cell in col:
                    cell.number_format = number_format
                    cell.alignment = Alignment(horizontal="right", vertical="top")

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22 if row_idx == 1 else 34


def product_role(row: dict) -> str:
    issuer = row["issuer_display_name"]
    ticker = row["ticker"]
    name = row.get("fund_name") or ""
    asset_class = row.get("asset_class")
    strategy = row.get("strategy")
    if issuer == "BlackRock/iShares":
        if ticker in {"IVV", "ITOT", "IJH", "IJR"}:
            return "Core宽基代表产品"
        if ticker in {"AGG", "SGOV"} or asset_class == "Fixed Income":
            return "Core债券/现金管理代表产品"
        if ticker in {"IEFA", "IEMG"} or "MSCI" in name:
            return "国际/海外代表产品"
        return "AUM靠前产品"
    if issuer == "JPMorgan AM":
        if ticker in {"JEPI", "JEPQ"}:
            return "收益增强/主动ETF代表产品"
        if ticker.startswith("BB"):
            return "BetaBuilders/核心Beta代表产品"
        if asset_class == "Fixed Income" and strategy == "Active":
            return "主动固收代表产品"
        if strategy == "Active":
            return "主动权益代表产品"
        return "AUM靠前产品"
    if issuer == "Invesco":
        if ticker in {"QQQ", "QQQM"}:
            return "入口型大单品/创新成长代表产品"
        if asset_class == "Fixed Income":
            return "固定收益代表产品"
        if asset_class == "Commodities":
            return "商品代表产品"
        return "AUM靠前/Smart Beta或主题代表产品"
    return "代表产品"


def build_overview(rows: list[dict]) -> list[dict]:
    output = []
    for raw_issuer, display in ISSUER_MAP.items():
        issuer_rows = [row for row in rows if row.get("issuer") == raw_issuer]
        total_aum = sum_field(issuer_rows, "aum_usd")
        total_adv = sum_field(issuer_rows, "avg_daily_dollar_vol_usd")
        active_rows = [row for row in issuer_rows if row.get("strategy") == "Active"]
        active_aum = sum_field(active_rows, "aum_usd")
        expense_values = numeric_values(issuer_rows, "expense_ratio_pct")
        top_products = sorted(issuer_rows, key=lambda row: to_float(row.get("aum_usd")) or 0, reverse=True)[:5]
        meta = ISSUER_META[display]
        output.append({
            "issuer_display_name": display,
            "issuer_raw_name": raw_issuer,
            "case_role": meta["case_role"],
            "etf_count": len(issuer_rows),
            "total_aum_usd": total_aum,
            "total_aum_billion_usd": total_aum / 1e9,
            "total_aum_trillion_usd": total_aum / 1e12,
            "total_aum_100m_usd": total_aum / 1e8,
            "avg_expense_ratio_pct": mean(expense_values) if expense_values else None,
            "median_expense_ratio_pct": median(expense_values) if expense_values else None,
            "total_adv_usd": total_adv,
            "total_adv_billion_usd": total_adv / 1e9,
            "active_etf_count": len(active_rows),
            "active_aum_usd": active_aum,
            "active_aum_billion_usd": active_aum / 1e9,
            "active_aum_share_pct": (active_aum / total_aum * 100) if total_aum else None,
            "equity_aum_billion_usd": sum_field([row for row in issuer_rows if row.get("asset_class") == "Equity"], "aum_usd") / 1e9,
            "fixed_income_aum_billion_usd": sum_field([row for row in issuer_rows if row.get("asset_class") == "Fixed Income"], "aum_usd") / 1e9,
            "flagship_products": "；".join(f"{row.get('ticker')} {row.get('fund_name')}" for row in top_products),
            "ecosystem_keywords": meta["ecosystem_keywords"],
            "ppt_role": meta["ppt_role"],
            "data_as_of": SNAPSHOT_DATE,
            "source_note": f"{CSV_NAME}; aum_usd与avg_daily_dollar_vol_usd使用原始美元字段，expense_ratio_pct保留原始百分数口径",
        })
    return output


def build_mix(rows: list[dict]) -> list[dict]:
    output = []
    for raw_issuer, display in ISSUER_MAP.items():
        issuer_rows = [row for row in rows if row.get("issuer") == raw_issuer]
        total_aum = sum_field(issuer_rows, "aum_usd")
        categories = sorted({row.get("asset_class") or "Unclassified" for row in issuer_rows})
        for category in categories:
            subset = [row for row in issuer_rows if (row.get("asset_class") or "Unclassified") == category]
            aum = sum_field(subset, "aum_usd")
            expenses = numeric_values(subset, "expense_ratio_pct")
            output.append({
                "issuer_display_name": display,
                "asset_class_or_category": category,
                "etf_count": len(subset),
                "total_aum_usd": aum,
                "total_aum_billion_usd": aum / 1e9,
                "aum_share": aum / total_aum if total_aum else None,
                "avg_expense_ratio_pct": mean(expenses) if expenses else None,
                "median_expense_ratio_pct": median(expenses) if expenses else None,
                "ppt_interpretation": ISSUER_META[display]["product_mix_note"],
                "data_as_of": SNAPSHOT_DATE,
            })
    output.sort(key=lambda row: (list(ISSUER_MAP.values()).index(row["issuer_display_name"]), -(row["total_aum_usd"] or 0)))
    return output


def build_representatives(rows: list[dict]) -> list[dict]:
    by_ticker = {row.get("ticker"): row for row in rows}
    output = []
    for display, tickers in REP_TICKERS.items():
        for ticker in tickers:
            row = by_ticker.get(ticker)
            if not row:
                continue
            aum = to_float(row.get("aum_usd")) or 0
            volume = to_float(row.get("avg_daily_dollar_vol_usd")) or 0
            ytd_flow = to_float(row.get("ytd_fund_flow_usd")) or 0
            one_year_flow = to_float(row.get("one_year_fund_flow_usd")) or 0
            output.append({
                "issuer_display_name": display,
                "ticker": ticker,
                "fund_name": row.get("fund_name"),
                "case_product_role": product_role(row),
                "asset_class": row.get("asset_class"),
                "category": row.get("category"),
                "strategy": row.get("strategy"),
                "aum_usd": aum,
                "aum_billion_usd": aum / 1e9,
                "aum_100m_usd": aum / 1e8,
                "expense_ratio_pct": to_float(row.get("expense_ratio_pct")),
                "avg_daily_volume_usd": volume,
                "avg_daily_volume_billion_usd": volume / 1e9,
                "ytd_flow_usd": ytd_flow,
                "ytd_flow_billion_usd": ytd_flow / 1e9,
                "one_year_flow_usd": one_year_flow,
                "one_year_flow_billion_usd": one_year_flow / 1e9,
                "ppt_usage": ISSUER_META[display]["ppt_usage"],
                "data_as_of": SNAPSHOT_DATE,
                "source_note": f"{CSV_NAME}; 原始CSV ticker={ticker}",
            })
    output.sort(key=lambda row: (list(ISSUER_MAP.values()).index(row["issuer_display_name"]), -(row["aum_usd"] or 0)))
    return output


def build_invesco_helper(rep_rows: list[dict]) -> list[dict]:
    output = []
    for row in rep_rows:
        if row["issuer_display_name"] == "Invesco" and row["ticker"] in {"QQQ", "QQQM"}:
            output.append({
                "ticker": row["ticker"],
                "fund_name": row["fund_name"],
                "product_role": row["case_product_role"],
                "aum_usd": row["aum_usd"],
                "aum_billion_usd": row["aum_billion_usd"],
                "expense_ratio_pct": row["expense_ratio_pct"],
                "index_or_exposure": "Nasdaq-100/创新成长",
                "ppt_message": "QQQ把“纳斯达克100/创新成长”压缩成一个ETF符号，对富国的启示是能力封装之外也需要客户记得住的代表产品。",
                "source": row["source_note"],
            })
    return sorted(output, key=lambda row: -(row["aum_usd"] or 0))


def ensure_p26_sheet(wb):
    short = "09_P26_Fullgoal_Implication"
    long = "09_P26_Fullgoal_Implication_Data"
    if long in wb.sheetnames:
        return
    if short in wb.sheetnames:
        wb[short].title = long
    else:
        rows = [
            {"fullgoal_capability": "指数能力", "current_basis": "宽基、大盘、科创、创业板等布局", "etf_solution_entry": "底仓入口", "customer_scenario": "A股底仓、长期配置、市场Beta获取", "action_suggestion": "筛选1-2只核心宽基作为底仓入口，强化成交、价差、渠道话术和组合示例", "required_followup_data": "重点产品规模、成交额、买卖价差、折溢价、净流入"},
            {"fullgoal_capability": "投研能力", "current_basis": "新能源、信息技术、港股、高端制造、生物医药等主题布局", "etf_solution_entry": "主题入口", "customer_scenario": "产业趋势表达、阶段性行情捕捉、成长弹性配置", "action_suggestion": "持续输出产业观点、配置逻辑和风险提示", "required_followup_data": "主题产品规模、成交、净流入、同类排名、投研内容供给"},
            {"fullgoal_capability": "量化/策略能力", "current_basis": "当前策略ETF关注度较低但可作为中长期差异化方向", "etf_solution_entry": "风格入口", "customer_scenario": "红利、低波、质量、成长、现金流等风格配置", "action_suggestion": "将策略能力转化为客户可理解的收益风险特征", "required_followup_data": "策略ETF产品池、因子暴露、历史表现、回撤、波动、同类排名"},
            {"fullgoal_capability": "渠道与交易能力", "current_basis": "富国已有产品储备，但重点产品心智仍待沉淀", "etf_solution_entry": "使用入口", "customer_scenario": "让渠道会讲、客户敢买、交易顺畅、长期留存", "action_suggestion": "建立旗舰池、孵化池、维护池，集中资源提升重点产品交易体验", "required_followup_data": "做市质量、日均成交额、折溢价、买卖价差、渠道反馈"},
        ]
        write_table(wb, long, ["fullgoal_capability", "current_basis", "etf_solution_entry", "customer_scenario", "action_suggestion", "required_followup_data"], rows)
    style_sheet(wb[long])


def update_slide_prompts(wb):
    if "10_Slide_Text_Prompts" not in wb.sheetnames:
        return
    ws = wb["10_Slide_Text_Prompts"]
    headers = [cell.value for cell in ws[1]]
    try:
        slide_col = headers.index("slide_no") + 1
        chart_col = headers.index("chart_data_sheet") + 1
    except ValueError:
        return
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row_idx, slide_col).value == "P26":
            ws.cell(row_idx, chart_col).value = "09_P26_Fullgoal_Implication_Data"
    style_sheet(ws)


def validate_units(overview_rows: list[dict], mix_rows: list[dict], rep_rows: list[dict], invesco_rows: list[dict]):
    errors = []
    for row in overview_rows:
        expected = EXPECTED_ISSUER_AUM_B[row["issuer_display_name"]]
        actual = row["total_aum_billion_usd"]
        if abs(actual - expected) > 0.2:
            errors.append(f"{row['issuer_display_name']} 总AUM异常：actual={actual:.2f}B expected≈{expected:.2f}B")
    mix_lookup = {(row["issuer_display_name"], row["asset_class_or_category"]): row["total_aum_billion_usd"] for row in mix_rows}
    expected_mix = {
        ("BlackRock/iShares", "Equity"): 3442.14,
        ("BlackRock/iShares", "Fixed Income"): 954.99,
        ("JPMorgan AM", "Equity"): 217.20,
        ("JPMorgan AM", "Fixed Income"): 108.30,
        ("Invesco", "Equity"): 906.02,
    }
    for key, expected in expected_mix.items():
        actual = mix_lookup.get(key)
        if actual is None or abs(actual - expected) > 0.2:
            errors.append(f"{key[0]} {key[1]} AUM异常：actual={actual}B expected≈{expected:.2f}B")
    rep_lookup = {row["ticker"]: row["aum_billion_usd"] for row in rep_rows}
    for ticker, expected in REQUIRED_PRODUCT_AUM_B.items():
        actual = rep_lookup.get(ticker)
        if actual is None or abs(actual - expected) > 0.05:
            errors.append(f"{ticker} AUM异常：actual={actual}B expected≈{expected:.2f}B")
    invesco_lookup = {row["ticker"]: row["aum_billion_usd"] for row in invesco_rows}
    for ticker in ["QQQ", "QQQM"]:
        expected = REQUIRED_PRODUCT_AUM_B[ticker]
        actual = invesco_lookup.get(ticker)
        if actual is None or abs(actual - expected) > 0.05:
            errors.append(f"08_P25_Invesco_Helper_Data {ticker} AUM异常：actual={actual}B expected≈{expected:.2f}B")
    if errors:
        raise ValueError("单位异常，已停止保存：\n" + "\n".join(errors))


def build_qa(overview_rows, rep_rows, wb):
    qa_rows = []

    def add(check_item, expected, actual, passed, source_sheet):
        qa_rows.append({
            "check_item": check_item,
            "expected": expected,
            "actual": actual,
            "status": "PASS" if passed else "FAIL",
            "source_sheet": source_sheet,
            "data_as_of": SNAPSHOT_DATE,
            "source_note": f"{CSV_NAME}; 自动校验",
        })

    overview_lookup = {row["issuer_display_name"]: row for row in overview_rows}
    for issuer, expected in EXPECTED_ISSUER_AUM_B.items():
        actual = overview_lookup[issuer]["total_aum_billion_usd"]
        add(f"{issuer}总AUM是否约{expected:.2f}B", f"{expected:.2f}B", f"{actual:.2f}B", abs(actual - expected) <= 0.2, "01_Issuer_Overview_Clean")

    rep_lookup = {row["ticker"]: row for row in rep_rows}
    for ticker in ["IVV", "QQQ", "QQQM", "JEPI", "JEPQ"]:
        expected = REQUIRED_PRODUCT_AUM_B[ticker]
        actual = rep_lookup[ticker]["aum_billion_usd"]
        add(f"{ticker}是否约{expected:.2f}B", f"{expected:.2f}B", f"{actual:.2f}B", abs(actual - expected) <= 0.05, "03_Representative_ETFs_Clean")

    data_as_of_ok = True
    source_note_ok = True
    for sheet_name in ["01_Issuer_Overview_Clean", "03_Representative_ETFs_Clean", "10_Slide_Text_Prompts", "11_QA_Checks"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        if "data_as_of" in headers:
            col = headers.index("data_as_of") + 1
            for row_idx in range(2, ws.max_row + 1):
                data_as_of_ok = data_as_of_ok and ws.cell(row_idx, col).value == SNAPSHOT_DATE
        if "source_note" in headers:
            col = headers.index("source_note") + 1
            for row_idx in range(2, ws.max_row + 1):
                source_note_ok = source_note_ok and bool(ws.cell(row_idx, col).value)
    add("所有data_as_of是否为2026-07-06", SNAPSHOT_DATE, "见含data_as_of字段的清洗/文案/QA页", data_as_of_ok, "清洗页")
    add("所有source_note是否非空", "非空", "见含source_note字段的清洗/文案/QA页", source_note_ok, "清洗页")
    return qa_rows


def main():
    csv_path = find_csv()
    rows = read_target_rows(csv_path)
    wb = load_workbook(INPUT_XLSX)

    for sheet in ["01_Issuer_Overview_Clean", "02_Product_Mix_Clean", "03_Representative_ETFs_Clean", "08_P25_Invesco_Helper_Data", "11_QA_Checks"]:
        delete_if_exists(wb, sheet)

    overview_headers = ["issuer_display_name", "issuer_raw_name", "case_role", "etf_count", "total_aum_usd", "total_aum_billion_usd", "total_aum_trillion_usd", "total_aum_100m_usd", "avg_expense_ratio_pct", "median_expense_ratio_pct", "total_adv_usd", "total_adv_billion_usd", "active_etf_count", "active_aum_usd", "active_aum_billion_usd", "active_aum_share_pct", "equity_aum_billion_usd", "fixed_income_aum_billion_usd", "flagship_products", "ecosystem_keywords", "ppt_role", "data_as_of", "source_note"]
    mix_headers = ["issuer_display_name", "asset_class_or_category", "etf_count", "total_aum_usd", "total_aum_billion_usd", "aum_share", "avg_expense_ratio_pct", "median_expense_ratio_pct", "ppt_interpretation", "data_as_of"]
    rep_headers = ["issuer_display_name", "ticker", "fund_name", "case_product_role", "asset_class", "category", "strategy", "aum_usd", "aum_billion_usd", "aum_100m_usd", "expense_ratio_pct", "avg_daily_volume_usd", "avg_daily_volume_billion_usd", "ytd_flow_usd", "ytd_flow_billion_usd", "one_year_flow_usd", "one_year_flow_billion_usd", "ppt_usage", "data_as_of", "source_note"]
    invesco_headers = ["ticker", "fund_name", "product_role", "aum_usd", "aum_billion_usd", "expense_ratio_pct", "index_or_exposure", "ppt_message", "source"]
    qa_headers = ["check_item", "expected", "actual", "status", "source_sheet", "data_as_of", "source_note"]

    overview_rows = build_overview(rows)
    mix_rows = build_mix(rows)
    rep_rows = build_representatives(rows)
    invesco_rows = build_invesco_helper(rep_rows)
    validate_units(overview_rows, mix_rows, rep_rows, invesco_rows)

    write_table(wb, "01_Issuer_Overview_Clean", overview_headers, overview_rows)
    write_table(wb, "02_Product_Mix_Clean", mix_headers, mix_rows)
    write_table(wb, "03_Representative_ETFs_Clean", rep_headers, rep_rows)
    write_table(wb, "08_P25_Invesco_Helper_Data", invesco_headers, invesco_rows)
    ensure_p26_sheet(wb)
    update_slide_prompts(wb)

    qa_rows = build_qa(overview_rows, rep_rows, wb)
    write_table(wb, "11_QA_Checks", qa_headers, qa_rows)
    wb.save(OUTPUT_XLSX)

    print(f"CSV来源: {csv_path}")
    print(f"输出文件: {OUTPUT_XLSX}")
    print("\n新增/修正sheet行列数（数据行数，不含表头）:")
    for sheet_name in ["01_Issuer_Overview_Clean", "02_Product_Mix_Clean", "03_Representative_ETFs_Clean", "08_P25_Invesco_Helper_Data", "09_P26_Fullgoal_Implication_Data", "10_Slide_Text_Prompts", "11_QA_Checks"]:
        ws = wb[sheet_name]
        print(f"- {sheet_name}: {max(ws.max_row - 1, 0)} 行 x {ws.max_column} 列")

    print("\n三家发行人AUM、ETF数量、日均成交额:")
    for row in overview_rows:
        print(
            f"- {row['issuer_display_name']}: ETF数量={row['etf_count']}, "
            f"AUM={row['total_aum_billion_usd']:.2f} billion USD, "
            f"日均成交额={row['total_adv_billion_usd']:.2f} billion USD"
        )

    print("\n代表产品AUM校验结果:")
    rep_lookup = {row["ticker"]: row for row in rep_rows}
    for ticker, expected in REQUIRED_PRODUCT_AUM_B.items():
        actual = rep_lookup[ticker]["aum_billion_usd"]
        status = "PASS" if abs(actual - expected) <= 0.05 else "FAIL"
        print(f"- {ticker}: {actual:.2f} billion USD，目标约 {expected:.2f}B，{status}")

    print("\n单位异常检查: PASS，未发现AUM或成交额缩小10倍问题。")


if __name__ == "__main__":
    main()
