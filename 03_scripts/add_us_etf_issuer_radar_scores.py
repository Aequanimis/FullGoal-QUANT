from __future__ import annotations

import math
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SNAPSHOT_DATE = "2026-07-06"

INPUT_PATH = (
    Path.home()
    / "Desktop"
    / "Fullgoal"
    / "ICI课题"
    / "课题更正"
    / "美国ETF案例数据收集_富国启示_作图sheet版.xlsx"
)
OUTPUT_PATH = (
    Path.home()
    / "Desktop"
    / "Fullgoal"
    / "ICI课题"
    / "课题更正"
    / "美国ETF案例数据收集_富国启示_作图sheet版_v2_含评分.xlsx"
)

ISSUERS = ["BlackRock/iShares", "JPMorgan AM", "Invesco"]
DIMENSIONS = [
    "产品矩阵广度",
    "规模沉淀能力",
    "成本竞争力",
    "交易流动性",
    "能力迁移程度",
    "入口产品心智",
]

NEW_SHEETS = [
    "12_US_Issuer_Radar_Score",
    "13_US_Issuer_Radar_Long",
    "14_US_Score_Methodology",
]


def to_float(value) -> float | None:
    if value in (None, ""):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(number) else number


def read_table(wb, sheet_name: str) -> list[dict]:
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[1]]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        row = {header: values[idx] for idx, header in enumerate(headers) if header}
        if any(value not in (None, "") for value in row.values()):
            rows.append(row)
    return rows


def normalize(raw_values: dict[str, float], positive: bool = True) -> dict[str, float]:
    values = list(raw_values.values())
    min_value = min(values)
    max_value = max(values)
    if max_value == min_value:
        return {issuer: 80.0 for issuer in raw_values}
    if positive:
        return {issuer: 60 + 40 * (value - min_value) / (max_value - min_value) for issuer, value in raw_values.items()}
    return {issuer: 60 + 40 * (max_value - value) / (max_value - min_value) for issuer, value in raw_values.items()}


def safe_log1p(value: float | None) -> float:
    return math.log(1 + max(value or 0, 0))


def safe_log(value: float | None) -> float:
    return math.log(max(value or 0, 1e-12))


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
        lower = header.lower()
        max_len = len(header)
        for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
        is_long_text = any(token in lower for token in ["note", "source", "definition", "variable", "limitation", "说明", "局限", "scoring"])
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 52 if is_long_text else 24)

        if lower.endswith("_score") or header == "score":
            for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                for cell in col:
                    cell.number_format = "0.0"
                    cell.alignment = Alignment(horizontal="right", vertical="top")
        elif any(token in lower for token in ["raw", "aum", "adv", "expense", "share", "count"]):
            for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                for cell in col:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0.00"

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22 if row_idx == 1 else 36


def write_table(wb, sheet_name: str, headers: list[str], rows: list[dict]):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])
    style_sheet(ws)
    return ws


def issuer_rows_by_name(rows: list[dict]) -> dict[str, dict]:
    return {row.get("issuer_display_name"): row for row in rows if row.get("issuer_display_name")}


def build_scores(wb):
    missing_fields: list[str] = []

    overview_rows = read_table(wb, "01_Issuer_Overview_Clean")
    product_mix_rows = read_table(wb, "02_Product_Mix_Clean")
    representative_rows = read_table(wb, "03_Representative_ETFs_Clean")
    official_rows = read_table(wb, "04_Official_Sources_Clean")
    invesco_rows = read_table(wb, "08_P25_Invesco_Helper_Data")
    qa_rows = read_table(wb, "11_QA_Checks")

    if not product_mix_rows:
        missing_fields.append("02_Product_Mix_Clean sheet不存在，产品矩阵广度改用03_Representative_ETFs_Clean的asset_class/category/strategy近似计算")
    if not official_rows:
        missing_fields.append("04_Official_Sources_Clean sheet不存在")
    if not qa_rows:
        missing_fields.append("11_QA_Checks sheet不存在")

    overview = issuer_rows_by_name(overview_rows)
    for issuer in ISSUERS:
        if issuer not in overview:
            missing_fields.append(f"01_Issuer_Overview_Clean缺少发行人：{issuer}")

    reps_by_issuer = {issuer: [row for row in representative_rows if row.get("issuer_display_name") == issuer] for issuer in ISSUERS}

    raw: dict[str, dict[str, float]] = {dimension: {} for dimension in DIMENSIONS}
    raw_detail: dict[str, dict[str, float]] = {issuer: {} for issuer in ISSUERS}

    for issuer in ISSUERS:
        if product_mix_rows:
            issuer_mix = [row for row in product_mix_rows if row.get("issuer_display_name") == issuer]
            asset_count = len({row.get("asset_class_or_category") for row in issuer_mix if row.get("asset_class_or_category")})
            extra_count = 0
        else:
            issuer_reps = reps_by_issuer[issuer]
            asset_count = len({row.get("asset_class") for row in issuer_reps if row.get("asset_class")})
            categories = {row.get("category") for row in issuer_reps if row.get("category")}
            strategies = {row.get("strategy") for row in issuer_reps if row.get("strategy")}
            extra_count = len(categories | strategies)
        product_breadth = asset_count + 0.5 * extra_count
        raw["产品矩阵广度"][issuer] = product_breadth
        raw_detail[issuer]["product_breadth_raw"] = product_breadth

        issuer_overview = overview.get(issuer, {})
        total_aum = to_float(issuer_overview.get("total_aum_billion_usd")) or 0
        large_product_count = sum(1 for row in reps_by_issuer[issuer] if (to_float(row.get("aum_billion_usd")) or 0) > 10)
        scale_raw = 0.6 * safe_log(total_aum) + 0.4 * safe_log1p(large_product_count)
        raw["规模沉淀能力"][issuer] = scale_raw
        raw_detail[issuer]["total_aum_billion_usd"] = total_aum
        raw_detail[issuer]["large_product_count"] = large_product_count
        raw_detail[issuer]["scale_raw"] = scale_raw

        expense = to_float(issuer_overview.get("avg_expense_ratio_pct"))
        expense_source = "avg_expense_ratio_pct"
        if expense is None:
            expense = to_float(issuer_overview.get("median_expense_ratio_pct"))
            expense_source = "median_expense_ratio_pct"
        if expense is None:
            expense = 0
            missing_fields.append(f"{issuer}缺少avg_expense_ratio_pct和median_expense_ratio_pct，成本竞争力暂按0处理")
        raw["成本竞争力"][issuer] = expense
        raw_detail[issuer]["expense_raw_pct"] = expense
        raw_detail[issuer]["expense_source"] = expense_source

        total_adv = to_float(issuer_overview.get("total_adv_billion_usd")) or 0
        liquidity_raw = safe_log1p(total_adv)
        raw["交易流动性"][issuer] = liquidity_raw
        raw_detail[issuer]["total_adv_billion_usd"] = total_adv
        raw_detail[issuer]["liquidity_raw"] = liquidity_raw

        active_aum_share = to_float(issuer_overview.get("active_aum_share_pct")) or 0
        active_count = to_float(issuer_overview.get("active_etf_count")) or 0
        fixed_income_aum = to_float(issuer_overview.get("fixed_income_aum_billion_usd"))
        if fixed_income_aum is None:
            fixed_income_aum = 0
            missing_fields.append(f"{issuer}缺少fixed_income_aum_billion_usd，能力迁移程度固收项按0处理")
        epi_aum = 0.0
        for row in reps_by_issuer[issuer]:
            ticker = str(row.get("ticker") or "")
            fund_name = str(row.get("fund_name") or "")
            if ticker in {"JEPI", "JEPQ"} or "Equity Premium Income" in fund_name or "Premium Income" in fund_name:
                epi_aum += to_float(row.get("aum_billion_usd")) or 0
        capability_raw = (
            0.4 * active_aum_share
            + 0.2 * safe_log1p(active_count)
            + 0.2 * safe_log1p(fixed_income_aum)
            + 0.2 * safe_log1p(epi_aum)
        )
        raw["能力迁移程度"][issuer] = capability_raw
        raw_detail[issuer]["active_aum_share_pct"] = active_aum_share
        raw_detail[issuer]["active_etf_count"] = active_count
        raw_detail[issuer]["fixed_income_aum_billion_usd"] = fixed_income_aum
        raw_detail[issuer]["equity_premium_income_aum_billion_usd"] = epi_aum
        raw_detail[issuer]["capability_migration_raw"] = capability_raw

        top_product_aum = max([to_float(row.get("aum_billion_usd")) or 0 for row in reps_by_issuer[issuer]] or [0])
        top_product_share = top_product_aum / total_aum if total_aum else 0
        flagship_raw = 0.7 * safe_log1p(top_product_aum) + 0.3 * top_product_share
        raw["入口产品心智"][issuer] = flagship_raw
        raw_detail[issuer]["top_product_aum_billion_usd"] = top_product_aum
        raw_detail[issuer]["top_product_share"] = top_product_share
        raw_detail[issuer]["flagship_raw"] = flagship_raw

    scores_by_dimension = {
        "产品矩阵广度": normalize(raw["产品矩阵广度"], positive=True),
        "规模沉淀能力": normalize(raw["规模沉淀能力"], positive=True),
        "成本竞争力": normalize(raw["成本竞争力"], positive=False),
        "交易流动性": normalize(raw["交易流动性"], positive=True),
        "能力迁移程度": normalize(raw["能力迁移程度"], positive=True),
        "入口产品心智": normalize(raw["入口产品心智"], positive=True),
    }

    wide_rows = []
    long_rows = []
    raw_rows = []
    for issuer in ISSUERS:
        notes = [
            "评分用于PPT展示三类ETF生态路径差异，不作为正式排名。",
            "分数按三家公司样本内60-100归一化。",
        ]
        if issuer == "Invesco" and any(row.get("ticker") == "QQQ" for row in invesco_rows + reps_by_issuer[issuer]):
            notes.append("QQQ具有显著入口型大单品心智，入口产品心智得分需结合定性判断解释。")
        if not product_mix_rows:
            notes.append("产品矩阵广度因轻量作图包缺少02_Product_Mix_Clean，使用代表产品sheet近似计算。")

        wide_row = {
            "issuer_display_name": issuer,
            "产品矩阵广度_score": scores_by_dimension["产品矩阵广度"][issuer],
            "规模沉淀能力_score": scores_by_dimension["规模沉淀能力"][issuer],
            "成本竞争力_score": scores_by_dimension["成本竞争力"][issuer],
            "交易流动性_score": scores_by_dimension["交易流动性"][issuer],
            "能力迁移程度_score": scores_by_dimension["能力迁移程度"][issuer],
            "入口产品心智_score": scores_by_dimension["入口产品心智"][issuer],
            "scoring_note": " ".join(notes),
            "data_as_of": SNAPSHOT_DATE,
            "source_note": "01_Issuer_Overview_Clean；03_Representative_ETFs_Clean；04_Official_Sources_Clean；08_P25_Invesco_Helper_Data；11_QA_Checks",
        }
        wide_rows.append(wide_row)

        for dimension in DIMENSIONS:
            long_rows.append({
                "issuer_display_name": issuer,
                "dimension": dimension,
                "score": scores_by_dimension[dimension][issuer],
                "data_as_of": SNAPSHOT_DATE,
                "scoring_note": wide_row["scoring_note"],
            })

        raw_rows.append({
            "issuer_display_name": issuer,
            **raw_detail[issuer],
        })

    methodology_rows = build_methodology_rows(product_mix_rows_exists=bool(product_mix_rows))
    return wide_rows, long_rows, methodology_rows, raw_rows, missing_fields


def build_methodology_rows(product_mix_rows_exists: bool) -> list[dict]:
    product_source = "02_Product_Mix_Clean中的asset_class_or_category数量" if product_mix_rows_exists else "03_Representative_ETFs_Clean中的asset_class/category/strategy近似值"
    rows = [
        {"section": "评分用途", "content": "用于PPT展示三类美国ETF生态路径差异，不作为正式排名", "data_as_of": SNAPSHOT_DATE, "source_note": "用户任务说明；研究判断"},
        {"section": "数据截至日期", "content": SNAPSHOT_DATE, "data_as_of": SNAPSHOT_DATE, "source_note": "用户指定"},
        {"section": "使用数据来源", "content": "境外ETF全量数据、发行人官网资料、研究判断", "data_as_of": SNAPSHOT_DATE, "source_note": "作图sheet版工作簿"},
        {"section": "维度定义", "content": "产品矩阵广度：发行人ETF产品覆盖资产类别、产品类别和策略的广度。", "data_as_of": SNAPSHOT_DATE, "source_note": product_source},
        {"section": "维度定义", "content": "规模沉淀能力：发行人ETF总AUM与大型产品数量反映的规模承载能力。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean；03_Representative_ETFs_Clean"},
        {"section": "维度定义", "content": "成本竞争力：发行人平均或中位费率越低，得分越高。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean"},
        {"section": "维度定义", "content": "交易流动性：发行人整体日均成交额反映二级市场交易承载能力。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean"},
        {"section": "维度定义", "content": "能力迁移程度：主动ETF、主动AUM占比、固收AUM和收益增强产品反映传统资管能力ETF化程度。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean；03_Representative_ETFs_Clean"},
        {"section": "维度定义", "content": "入口产品心智：头部单品AUM与总AUM占比反映是否具备可代表发行人的ETF品牌入口。", "data_as_of": SNAPSHOT_DATE, "source_note": "03_Representative_ETFs_Clean；01_Issuer_Overview_Clean"},
        {"section": "原始变量", "content": "产品矩阵广度：product_breadth_raw = distinct asset_class_or_category count + 0.5 * distinct category/strategy count；若缺少02页则用代表产品表近似。", "data_as_of": SNAPSHOT_DATE, "source_note": product_source},
        {"section": "原始变量", "content": "规模沉淀能力：scale_raw = 0.6 * log(total_aum_billion_usd) + 0.4 * log(1 + large_product_count)。", "data_as_of": SNAPSHOT_DATE, "source_note": "01/03清洗页"},
        {"section": "原始变量", "content": "成本竞争力：expense_raw = avg_expense_ratio_pct；若缺失则用median_expense_ratio_pct，反向归一化。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean"},
        {"section": "原始变量", "content": "交易流动性：liquidity_raw = log(1 + total_adv_billion_usd)；当前作图包无median spread字段，未纳入spread项。", "data_as_of": SNAPSHOT_DATE, "source_note": "01_Issuer_Overview_Clean"},
        {"section": "原始变量", "content": "能力迁移程度：0.4 * active_aum_share_pct + 0.2 * log(1 + active_etf_count) + 0.2 * log(1 + fixed_income_aum_billion_usd) + 0.2 * log(1 + equity_premium_income_aum_billion_usd)。", "data_as_of": SNAPSHOT_DATE, "source_note": "01/03清洗页"},
        {"section": "原始变量", "content": "入口产品心智：flagship_raw = 0.7 * log(1 + top_product_aum_billion_usd) + 0.3 * top_product_share。", "data_as_of": SNAPSHOT_DATE, "source_note": "01/03清洗页"},
        {"section": "归一化方法", "content": "正向指标：score = 60 + 40 * (x - min_x) / (max_x - min_x)。", "data_as_of": SNAPSHOT_DATE, "source_note": "用户任务说明"},
        {"section": "归一化方法", "content": "反向指标：score = 60 + 40 * (max_x - x) / (max_x - min_x)。如max_x = min_x，则三家公司该项均记为80分。", "data_as_of": SNAPSHOT_DATE, "source_note": "用户任务说明"},
        {"section": "局限性", "content": "当前数据集缺少严格跟踪误差数据，故未设置“跟踪精度”维度。", "data_as_of": SNAPSHOT_DATE, "source_note": "研究判断"},
        {"section": "局限性", "content": "当前数据集缺少发行规模，故未完全复刻国内PPT中的“规模增长”维度。", "data_as_of": SNAPSHOT_DATE, "source_note": "研究判断"},
        {"section": "局限性", "content": "渠道生态和模型组合能力难以完全量化，仅通过官网资料和产品结构间接体现。", "data_as_of": SNAPSHOT_DATE, "source_note": "研究判断"},
        {"section": "局限性", "content": "样本只有三家公司，评分主要用于展示路径差异，不适合作为行业排名。", "data_as_of": SNAPSHOT_DATE, "source_note": "研究判断"},
    ]
    return rows


def validate_new_sheets(wb):
    missing = []
    for sheet_name in NEW_SHEETS:
        if sheet_name not in wb.sheetnames:
            missing.append(sheet_name)
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        for required in ["data_as_of"]:
            if required in headers:
                col = headers.index(required) + 1
                for row_idx in range(2, ws.max_row + 1):
                    if ws.cell(row_idx, col).value in (None, ""):
                        missing.append(f"{sheet_name}!{required} 第{row_idx}行为空")
        for required in ["source_note"]:
            if required in headers:
                col = headers.index(required) + 1
                for row_idx in range(2, ws.max_row + 1):
                    if ws.cell(row_idx, col).value in (None, ""):
                        missing.append(f"{sheet_name}!{required} 第{row_idx}行为空")
    return missing


def main():
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"输入文件不存在：{INPUT_PATH}")

    wb = load_workbook(INPUT_PATH)
    wide_rows, long_rows, methodology_rows, raw_rows, missing_fields = build_scores(wb)

    wide_headers = [
        "issuer_display_name",
        "产品矩阵广度_score",
        "规模沉淀能力_score",
        "成本竞争力_score",
        "交易流动性_score",
        "能力迁移程度_score",
        "入口产品心智_score",
        "scoring_note",
        "data_as_of",
        "source_note",
    ]
    long_headers = ["issuer_display_name", "dimension", "score", "data_as_of", "scoring_note"]
    methodology_headers = ["section", "content", "data_as_of", "source_note"]

    write_table(wb, "12_US_Issuer_Radar_Score", wide_headers, wide_rows)
    write_table(wb, "13_US_Issuer_Radar_Long", long_headers, long_rows)
    write_table(wb, "14_US_Score_Methodology", methodology_headers, methodology_rows)

    validation_missing = validate_new_sheets(wb)
    wb.save(OUTPUT_PATH)

    print("三家公司六维评分结果:")
    for row in wide_rows:
        print(
            f"- {row['issuer_display_name']}: "
            f"产品矩阵广度={row['产品矩阵广度_score']:.1f}, "
            f"规模沉淀能力={row['规模沉淀能力_score']:.1f}, "
            f"成本竞争力={row['成本竞争力_score']:.1f}, "
            f"交易流动性={row['交易流动性_score']:.1f}, "
            f"能力迁移程度={row['能力迁移程度_score']:.1f}, "
            f"入口产品心智={row['入口产品心智_score']:.1f}"
        )

    print("\n每个维度的原始指标值:")
    for row in raw_rows:
        print(
            f"- {row['issuer_display_name']}: "
            f"product_breadth_raw={row['product_breadth_raw']:.2f}, "
            f"scale_raw={row['scale_raw']:.4f}, "
            f"expense_raw_pct={row['expense_raw_pct']:.4f}, "
            f"liquidity_raw={row['liquidity_raw']:.4f}, "
            f"capability_migration_raw={row['capability_migration_raw']:.4f}, "
            f"flagship_raw={row['flagship_raw']:.4f}; "
            f"total_aum_billion_usd={row['total_aum_billion_usd']:.2f}, "
            f"large_product_count={row['large_product_count']}, "
            f"active_aum_share_pct={row['active_aum_share_pct']:.2f}, "
            f"active_etf_count={row['active_etf_count']:.0f}, "
            f"fixed_income_aum_billion_usd={row['fixed_income_aum_billion_usd']:.2f}, "
            f"equity_premium_income_aum_billion_usd={row['equity_premium_income_aum_billion_usd']:.2f}, "
            f"top_product_aum_billion_usd={row['top_product_aum_billion_usd']:.2f}, "
            f"top_product_share={row['top_product_share']:.4f}"
        )

    all_missing = missing_fields + validation_missing
    print("\n是否存在缺失字段:")
    if all_missing:
        print("是")
        for item in all_missing:
            print(f"- {item}")
    else:
        print("否")

    print("\n是否成功生成12、13、14三个新增sheet:")
    generated = all(sheet in wb.sheetnames for sheet in NEW_SHEETS)
    print("是" if generated else "否")
    print(f"\n输出文件: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
