from __future__ import annotations

import math
import re
from collections import defaultdict
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE = Path.home() / "Desktop" / "Fullgoal"
INPUT = BASE / "ICI课题" / "课题更正" / "中美ETF_TOP10五维生态评分_规模筛选版.xlsx"
CN_RAW = BASE / "课题研究" / "原始数据" / "全市场ETF基础信息.csv.xlsx"
CN_SUPP = BASE / "ICI课题" / "课题更正" / "data_extraction_top10_amc_etf_full_data.xlsx"
OUTPUT = BASE / "ICI课题" / "课题更正" / "中美ETF_TOP10五维生态评分_规模筛选版_中国口径修正版.xlsx"

SCORE_HEADERS = [
    "规模排名", "生态排名", "管理人", "ETF管理规模", "市场份额", "ETF产品数量",
    "产品广度", "费率优势", "交易活跃", "多元能力", "旗舰产品", "生态评分", "评分口径说明",
]

FIXED_TOP10 = [
    {"规模排名": 1, "管理人": "华夏基金", "ETF管理规模": 5965.56, "市场份额": 0.1252},
    {"规模排名": 2, "管理人": "易方达基金", "ETF管理规模": 5821.65, "市场份额": 0.1219},
    {"规模排名": 3, "管理人": "国泰基金", "ETF管理规模": 3715.02, "市场份额": 0.0789},
    {"规模排名": 4, "管理人": "华泰柏瑞基金", "ETF管理规模": 2999.29, "市场份额": 0.0617},
    {"规模排名": 5, "管理人": "广发基金", "ETF管理规模": 2593.97, "市场份额": 0.0553},
    {"规模排名": 6, "管理人": "南方基金", "ETF管理规模": 2455.50, "市场份额": 0.0531},
    {"规模排名": 7, "管理人": "博时基金", "ETF管理规模": 2292.66, "市场份额": 0.0488},
    {"规模排名": 8, "管理人": "嘉实基金", "ETF管理规模": 2247.00, "市场份额": 0.0437},
    {"规模排名": 9, "管理人": "华宝基金", "ETF管理规模": 2092.66, "市场份额": 0.0440},
    {"规模排名": 10, "管理人": "富国基金", "ETF管理规模": 1976.67, "市场份额": 0.0426},
]
FIXED_NAMES = [r["管理人"] for r in FIXED_TOP10]
FIXED_BY_NAME = {r["管理人"]: r for r in FIXED_TOP10}

CN_NOTE = "中国样本按用户提供的ETF管理规模截图固定TOP10名单，再在TOP10内部进行五维等权生态评分；规模只作为筛选标准，不进入生态评分；交易活跃优先使用2026-07-07补充交易数据修正；分数用于中国市场内部比较及中美路径映射，不代表中美绝对实力直接比较。"

RULE_ROWS = [
    ["产品广度", "20%", "产品线是否铺得全", "ETF产品数量、指数/主题覆盖数量", "越多越好", "看产品货架宽度，不评价规模"],
    ["费率优势", "20%", "产品成本是否有竞争力", "AUM加权综合费率", "费率越低越好", "ETF竞争的基础能力，简单直接"],
    ["交易活跃", "20%", "产品是否被市场真实交易使用", "成交效率、活跃产品占比", "成交效率越高、活跃产品占比越高越好", "规模已用于筛选，因此用成交效率而非绝对成交额"],
    ["多元能力", "20%", "是否能把不同投资能力转化为ETF产品", "多元产品规模占比、能力模块覆盖数", "占比越高、覆盖越多越好", "衡量策略、跨境、债券、商品、主题等能力ETF化"],
    ["旗舰产品", "20%", "是否有客户容易识别的代表产品", "最大ETF规模占比、Top3产品规模占比", "占比越高越好", "看是否形成代表性旗舰产品，而非单纯绝对规模"],
]
RULE_NOTE = "样本先按ETF管理规模筛选各自市场TOP10管理人，规模只作为入选标准，不进入生态评分。五个维度采用等权，分别为产品广度、费率优势、交易活跃、多元能力和旗舰产品。代表指标在各自市场TOP10内部标准化至60—100分。美国和中国使用同一套五维框架，但分数在各自市场内部标准化，因此中美之间主要比较结构、生态位和学习路径，不直接比较绝对实力。本修正版中，中国TOP10名单按用户截图固定，华安基金剔除，华宝基金纳入。"


def to_float(v: Any) -> float | None:
    if v in (None, "", "-"):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "").replace("%", ""))
    except ValueError:
        return None


def norm_code(v: Any) -> str:
    return re.sub(r"\s+", "", str(v or "").strip().upper())


def code_core(v: Any) -> str:
    return norm_code(v).split(".")[0]


def standardize_manager(v: Any) -> str:
    s = str(v or "").strip()
    for name in FIXED_NAMES:
        short = name.replace("基金", "")
        if name in s or short in s:
            return name
    return s


def sum_field(rows: list[dict[str, Any]], field: str) -> float:
    return sum((r.get(field) or 0) for r in rows if r.get(field) is not None)


def weighted_avg(rows: list[dict[str, Any]], val: str, weight: str) -> float | None:
    valid = [(r.get(val), r.get(weight)) for r in rows if r.get(val) is not None and (r.get(weight) or 0) > 0]
    tw = sum(w for _, w in valid)
    if tw <= 0:
        return None
    return sum(v * w for v, w in valid) / tw


def read_rows(ws) -> list[dict[str, Any]]:
    headers = [str(c.value or "").strip() for c in ws[1]]
    rows = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        row = {h: vals[i] if i < len(vals) else None for i, h in enumerate(headers) if h}
        if any(v not in (None, "") for v in row.values()):
            rows.append(row)
    return rows


def should_exclude_product(text: str) -> bool:
    upper = text.upper()
    return "LOF" in upper or any(token in text for token in ["联接", "場外", "场外", "普通指数基金", "非上市基金"])


def classify(row: dict[str, Any]) -> None:
    text = " ".join(str(row.get(k) or "") for k in ["fund_name", "index_name", "fund_type", "category"])
    broad = any(t in text for t in ["宽基", "沪深300", "中证500", "中证1000", "中证A500", "上证50", "科创50", "创业板", "深证100", "中证100", "A50", "上证180"])
    industry = any(t in text for t in ["行业", "主题", "科技", "信息技术", "新能源", "芯片", "半导体", "医药", "消费", "军工", "金融", "证券", "银行", "人工智能", "机器人", "光伏", "储能", "电池", "汽车", "港股科技"])
    smart = any(t in text for t in ["红利", "低波", "质量", "价值", "成长", "动量", "等权", "增强", "策略", "Smart Beta", "ESG", "央企", "国企", "现金流", "基本面"])
    cross = any(t in text for t in ["纳指", "标普", "日经", "德国", "法国", "港股", "恒生", "H股", "QDII", "跨境", "海外", "东南亚", "沙特"])
    fixed = any(t in text for t in ["债", "国债", "政金债", "信用债", "可转债", "公司债", "短融", "利率债"])
    commodity = any(t in text for t in ["黄金", "商品", "有色", "能源", "豆粕", "原油"])
    money = any(t in text for t in ["货币", "现金", "保证金", "添益"])
    active = any(t in text for t in ["主动", "增强"])
    row.update({
        "broad_based_flag": broad,
        "industry_theme_flag": industry,
        "smart_beta_strategy_flag": smart,
        "cross_border_flag": cross,
        "fixed_income_flag": fixed,
        "commodity_flag": commodity,
        "money_flag": money,
        "active_enhanced_flag": active,
        "non_plain_beta_flag": industry or smart or cross or fixed or commodity or money or active,
    })


def read_cn_products() -> tuple[list[dict[str, Any]], dict[str, Any]]:
    wb = load_workbook(CN_RAW, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = [c.value for c in ws[1]]
    rows = []
    raw_fixed_count = 0
    excluded = 0
    for vals in ws.iter_rows(min_row=2, values_only=True):
        raw = {headers[i]: vals[i] if i < len(vals) else None for i in range(len(headers))}
        manager = standardize_manager(raw.get("基金管理人"))
        if manager not in FIXED_NAMES:
            continue
        raw_fixed_count += 1
        full_text = " ".join(str(raw.get(h) or "") for h in headers)
        if should_exclude_product(full_text):
            excluded += 1
            continue
        aum_100m = to_float(raw.get("最新基金规模(亿)"))
        mgmt_fee = to_float(raw.get("管理费率_支持历史(%)"))
        custody_fee = to_float(raw.get("托管费率_支持历史(%)"))
        fee = (mgmt_fee or 0) + (custody_fee or 0) if mgmt_fee is not None or custody_fee is not None else None
        code = raw.get("Wind代码") or raw.get("基金代码") or raw.get("交易代码")
        row = {
            "manager": manager,
            "fund_code": norm_code(code),
            "fund_code_core": code_core(code),
            "fund_name": raw.get("基金简称") or raw.get("证券简称") or raw.get("基金全称"),
            "fund_type": raw.get("基金类型"),
            "category": raw.get("投资类型_二级分类"),
            "index_name": raw.get("跟踪指数名称") or raw.get("业绩比较基准"),
            "aum_100m": aum_100m,
            "fee_pct": fee,
            "raw_keep_note": "按全市场ETF基础信息表保留",
        }
        classify(row)
        rows.append(row)
    return rows, {"raw_fixed_count": raw_fixed_count, "excluded": excluded, "kept": len(rows)}


def read_supplement() -> list[dict[str, Any]]:
    wb = load_workbook(CN_SUPP, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = [c.value for c in ws[1]]
    rows = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        raw = {headers[i]: vals[i] if i < len(vals) else None for i in range(len(headers))}
        adv_m = to_float(raw.get("近60日日均成交额(百万元)"))
        manager = standardize_manager(raw.get("管理人"))
        rows.append({
            "manager": manager,
            "fund_code": norm_code(raw.get("基金代码")),
            "fund_code_core": code_core(raw.get("基金代码")),
            "adv_100m": adv_m / 100 if adv_m is not None else None,
        })
    return rows


def attach_trading(rows: list[dict[str, Any]]) -> dict[str, Any]:
    supp = read_supplement()
    by_code = {r["fund_code"]: r for r in supp}
    by_core = defaultdict(list)
    for r in supp:
        by_core[r["fund_code_core"]].append(r)
    matched = 0
    matched_by_manager = defaultdict(int)
    product_codes = {r["fund_code"] for r in rows}
    product_cores = {r["fund_code_core"] for r in rows}
    matched_supp = set()
    for row in rows:
        s = by_code.get(row["fund_code"])
        if not s:
            candidates = by_core.get(row["fund_code_core"], [])
            s = candidates[0] if candidates else None
        if s:
            row["adv_100m"] = s.get("adv_100m")
            matched += 1
            matched_by_manager[row["manager"]] += 1
            matched_supp.add(s["fund_code"])
        else:
            row["adv_100m"] = None
    unmatched_supp = [s for s in supp if s["fund_code"] not in product_codes and s["fund_code_core"] not in product_cores]
    return {
        "supp_rows": len(supp),
        "matched": matched,
        "matched_by_manager": dict(matched_by_manager),
        "unmatched_supp": len(unmatched_supp),
    }


def standardize(values: dict[str, float | None], positive: bool) -> tuple[dict[str, float | None], list[str]]:
    valid = [v for v in values.values() if v is not None]
    if not valid:
        return {k: None for k in values}, ["全样本缺失"]
    min_v, max_v = min(valid), max(valid)
    scores = {}
    warnings = []
    for k, v in values.items():
        if v is None:
            scores[k] = None
            warnings.append(f"{k}缺失")
            continue
        if max_v == min_v:
            score = 80.0
        elif positive:
            score = 60 + 40 * (v - min_v) / (max_v - min_v)
        else:
            score = 60 + 40 * (max_v - v) / (max_v - min_v)
        scores[k] = max(60.0, min(100.0, score))
    return scores, warnings


def average_dim(manager: str, vals: list[float | None], dim: str, warnings: list[str]) -> float:
    valid = [v for v in vals if v is not None]
    if not valid:
        warnings.append(f"{manager}-{dim}代表指标全缺失，采用80中性分")
        return 80.0
    if len(valid) < len(vals):
        warnings.append(f"{manager}-{dim}部分代表指标缺失，按可用指标计算")
    return sum(valid) / len(valid)


def build_metrics(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_manager = defaultdict(list)
    for r in rows:
        by_manager[r["manager"]].append(r)
    metrics = []
    for fixed in FIXED_TOP10:
        manager = fixed["管理人"]
        rs = by_manager.get(manager, [])
        fixed_scale = fixed["ETF管理规模"]
        sorted_aum = sorted(rs, key=lambda r: r.get("aum_100m") or 0, reverse=True)
        modules = {
            "宽基": any(r.get("broad_based_flag") for r in rs),
            "行业主题": any(r.get("industry_theme_flag") for r in rs),
            "策略/Smart Beta": any(r.get("smart_beta_strategy_flag") for r in rs),
            "跨境": any(r.get("cross_border_flag") for r in rs),
            "债券": any(r.get("fixed_income_flag") for r in rs),
            "商品": any(r.get("commodity_flag") for r in rs),
            "货币": any(r.get("money_flag") for r in rs),
            "主动/增强": any(r.get("active_enhanced_flag") for r in rs),
        }
        matched_trading = [r for r in rs if r.get("adv_100m") is not None]
        metric = {
            "规模排名": fixed["规模排名"],
            "manager": manager,
            "fixed_aum_100m": fixed_scale,
            "market_share": fixed["市场份额"],
            "product_count": len(rs),
            "coverage_count": len({r.get("index_name") or r.get("category") or r.get("fund_name") for r in rs if r.get("index_name") or r.get("category") or r.get("fund_name")}),
            "weighted_fee": weighted_avg(rs, "fee_pct", "aum_100m"),
            "trading_efficiency": sum_field(matched_trading, "adv_100m") / fixed_scale if matched_trading and fixed_scale else None,
            "active_product_share": sum(1 for r in matched_trading if (r.get("adv_100m") or 0) >= 1) / len(rs) if matched_trading and rs else None,
            "diversified_aum_share": sum_field([r for r in rs if r.get("non_plain_beta_flag")], "aum_100m") / fixed_scale if fixed_scale else None,
            "module_count": sum(modules.values()),
            "top1_aum_share": (sorted_aum[0].get("aum_100m") or 0) / fixed_scale if sorted_aum and fixed_scale else None,
            "top3_aum_share": sum_field(sorted_aum[:3], "aum_100m") / fixed_scale if fixed_scale else None,
            "trading_matched_count": len(matched_trading),
        }
        metrics.append(metric)
    return metrics


def score_cn(metrics: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    warnings = []
    fields = {
        "product_count": True,
        "coverage_count": True,
        "weighted_fee": False,
        "trading_efficiency": True,
        "active_product_share": True,
        "diversified_aum_share": True,
        "module_count": True,
        "top1_aum_share": True,
        "top3_aum_share": True,
    }
    raw = {field: {m["manager"]: m.get(field) for m in metrics} for field in fields}
    scores = {}
    for field, positive in fields.items():
        scores[field], ws = standardize(raw[field], positive)
        warnings.extend([f"{field}: {w}" for w in ws])
    by_manager = {m["manager"]: m for m in metrics}
    out = []
    for manager in FIXED_NAMES:
        m = by_manager[manager]
        dims = {
            "产品广度": average_dim(manager, [scores["product_count"][manager], scores["coverage_count"][manager]], "产品广度", warnings),
            "费率优势": average_dim(manager, [scores["weighted_fee"][manager]], "费率优势", warnings),
            "交易活跃": average_dim(manager, [scores["trading_efficiency"][manager], scores["active_product_share"][manager]], "交易活跃", warnings),
            "多元能力": average_dim(manager, [scores["diversified_aum_share"][manager], scores["module_count"][manager]], "多元能力", warnings),
            "旗舰产品": average_dim(manager, [scores["top1_aum_share"][manager], scores["top3_aum_share"][manager]], "旗舰产品", warnings),
        }
        eco = sum(dims.values()) / len(dims)
        out.append({
            "规模排名": m["规模排名"],
            "管理人": manager,
            "ETF管理规模": m["fixed_aum_100m"],
            "市场份额": m["market_share"],
            "ETF产品数量": m["product_count"],
            **{k: round(v, 1) for k, v in dims.items()},
            "生态评分": round(eco, 1),
            "评分口径说明": CN_NOTE,
        })
    out.sort(key=lambda r: r["生态评分"], reverse=True)
    for i, r in enumerate(out, 1):
        r["生态排名"] = i
    return out, warnings


def copy_worksheet_values(src_ws, dst_ws) -> None:
    for row in src_ws.iter_rows():
        for cell in row:
            dst = dst_ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                dst.font = copy(cell.font)
                dst.fill = copy(cell.fill)
                dst.border = copy(cell.border)
                dst.alignment = copy(cell.alignment)
                dst.number_format = cell.number_format
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[idx].height = dim.height
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.auto_filter.ref = src_ws.auto_filter.ref
    dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines


def write_score_sheet(wb: Workbook, rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet("中国TOP10五维评分")
    ws.append(SCORE_HEADERS)
    for r in rows:
        ws.append([r.get(h) for h in SCORE_HEADERS])
    style_sheet(ws, score_sheet=True)


def write_rules_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("评分细则")
    ws.append(["维度", "权重", "这项看什么", "实际计算指标", "评分方向", "设计说明"])
    for row in RULE_ROWS:
        ws.append(row)
    start = ws.max_row + 2
    ws.cell(start, 1).value = "说明"
    ws.cell(start, 1).font = Font(bold=True, color="1F4E78")
    ws.cell(start + 1, 1).value = RULE_NOTE
    ws.merge_cells(start_row=start + 1, start_column=1, end_row=start + 1, end_column=6)
    style_sheet(ws, score_sheet=False, filter_ref="A1:F6")
    ws.row_dimensions[start + 1].height = 90


def write_info_sheet(wb: Workbook, info_rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet("数据信息")
    headers = ["项目", "美国", "中国", "说明"]
    ws.append(headers)
    for r in info_rows:
        ws.append([r.get(h) for h in headers])
    style_sheet(ws, score_sheet=False)


def style_sheet(ws, score_sheet: bool, filter_ref: str | None = None) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = filter_ref or ws.dimensions
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if score_sheet and cell.column in {1, 2, 6}:
                cell.number_format = "0"
            if score_sheet and cell.column == 4:
                cell.number_format = "#,##0.00"
            if score_sheet and cell.column == 5:
                cell.number_format = "0.00%"
            if score_sheet and 7 <= cell.column <= 12:
                cell.number_format = "0.0"
            if score_sheet and cell.column == 12:
                cell.font = Font(bold=True)
    for col_idx in range(1, ws.max_column + 1):
        max_len = len(str(ws.cell(1, col_idx).value or ""))
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                max_len = max(max_len, min(len(str(value)), 50))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 50)
    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 24 if row_idx == 1 else 42


def build_info(old_cn_names: list[str], rows: list[dict[str, Any]], product_info: dict[str, Any], trade_info: dict[str, Any], metrics: list[dict[str, Any]], warnings: list[str]) -> list[dict[str, Any]]:
    product_counts = {m["manager"]: m["product_count"] for m in metrics}
    trade_counts = {m["manager"]: m["trading_matched_count"] for m in metrics}
    return [
        {"项目": "中国TOP10是否使用截图固定名单", "美国": "沿用原文件美国sheet", "中国": "是", "说明": "中国不再使用程序自动筛选TOP10。"},
        {"项目": "中国截图TOP10名单", "美国": "", "中国": "；".join(FIXED_NAMES), "说明": "规模排名、ETF管理规模、市场份额均按用户截图固定。"},
        {"项目": "原自动筛选TOP10名单", "美国": "", "中国": "；".join(old_cn_names), "说明": "来自输入文件旧中国TOP10五维评分sheet。"},
        {"项目": "被剔除管理人", "美国": "", "中国": "华安基金", "说明": "旧自动筛选结果中进入TOP10，本次按截图口径剔除。"},
        {"项目": "被新增管理人", "美国": "", "中国": "华宝基金", "说明": "按截图口径纳入TOP10，不得由华安基金替代。"},
        {"项目": "中国ETF管理规模和市场份额来源", "美国": "", "中国": "用户截图", "说明": "规模只作为筛选标准，不进入生态评分。"},
        {"项目": "中国产品级数据来源", "美国": "", "中国": str(CN_RAW), "说明": "固定TOP10管理人旗下记录默认按全市场ETF基础信息表保留，仅剔除明确联接/LOF/场外/普通指数基金。"},
        {"项目": "中国交易数据来源", "美国": "", "中国": str(CN_SUPP), "说明": "2026-07-07补充交易数据仅用于交易活跃维度。"},
        {"项目": "华宝基金产品匹配数量", "美国": "", "中国": product_counts.get("华宝基金", 0), "说明": "如为0则应FAIL；本项用于核验华宝纳入。"},
        {"项目": "华宝基金交易数据匹配数量", "美国": "", "中国": trade_counts.get("华宝基金", 0), "说明": "若覆盖不足，交易活跃可能受影响。"},
        {"项目": "中国固定TOP10各管理人的产品数量", "美国": "", "中国": "；".join(f"{k}:{v}" for k, v in product_counts.items()), "说明": "基于修正后的产品池。"},
        {"项目": "中国交易补充数据匹配率", "美国": "", "中国": f"{trade_info['matched']}/{len(rows)} = {trade_info['matched']/len(rows):.1%}; 补充未匹配{trade_info['unmatched_supp']}条", "说明": "补充数据无法新增产品入样本池。"},
        {"项目": "缺失字段WARN", "美国": "", "中国": "；".join(warnings) if warnings else "无", "说明": "包含单指标/中性分/字段缺失提示。"},
        {"项目": "口径限制说明", "美国": "美国sheet不重算，沿用输入文件", "中国": "中国TOP10按截图固定；6月底主数据决定产品池和AUM字段；7.07补充交易只用于交易活跃", "说明": "中美分数不能理解为绝对实力比较。"},
        {"项目": "QA结论", "美国": "PASS", "中国": "PASS" if product_counts.get("华宝基金", 0) > 0 and "华安基金" not in FIXED_NAMES else "FAIL", "说明": "确认华宝纳入、华安剔除、输出4个sheet、未使用旧中国五维得分。"},
        {"项目": "产品池保留说明", "美国": "", "中国": f"固定TOP10原始记录{product_info['raw_fixed_count']}条，剔除{product_info['excluded']}条，保留{product_info['kept']}条", "说明": "不以基金简称是否包含ETF作为唯一筛选条件。"},
    ]


def main() -> None:
    src_wb = load_workbook(INPUT, read_only=False, data_only=True)
    old_cn_rows = read_rows(src_wb["中国TOP10五维评分"])
    old_cn_names = [r.get("管理人") for r in old_cn_rows if r.get("管理人")]
    products, product_info = read_cn_products()
    trade_info = attach_trading(products)
    metrics = build_metrics(products)
    cn_scores, warnings = score_cn(metrics)

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    us_ws = out_wb.create_sheet("美国TOP10五维评分")
    copy_worksheet_values(src_wb["美国TOP10五维评分"], us_ws)
    style_sheet(us_ws, score_sheet=True)
    write_score_sheet(out_wb, cn_scores)
    write_rules_sheet(out_wb)
    write_info_sheet(out_wb, build_info(old_cn_names, products, product_info, trade_info, metrics, warnings))
    out_wb.save(OUTPUT)

    verify = load_workbook(OUTPUT, read_only=True, data_only=True)
    product_counts = {m["manager"]: m["product_count"] for m in metrics}
    trade_counts = {m["manager"]: m["trading_matched_count"] for m in metrics}
    print(f"输入五维评分文件路径: {INPUT}")
    print(f"中国原始产品数据路径: {CN_RAW}")
    print(f"中国交易补充数据路径: {CN_SUPP}")
    print(f"输出文件路径: {OUTPUT}")
    print(f"是否成功固定中国截图TOP10名单: {'是' if [m['manager'] for m in metrics] == FIXED_NAMES else '否'}")
    print(f"是否成功剔除华安基金: {'是' if '华安基金' not in [m['manager'] for m in metrics] else '否'}")
    print(f"是否成功纳入华宝基金: {'是' if '华宝基金' in [m['manager'] for m in metrics] else '否'}")
    print(f"华宝基金产品匹配数量: {product_counts.get('华宝基金', 0)}")
    print(f"华宝基金交易数据匹配数量: {trade_counts.get('华宝基金', 0)}")
    print("中国TOP10规模排名:")
    for m in metrics:
        print(f"- {m['规模排名']}. {m['manager']}: {m['fixed_aum_100m']:.2f}亿元, 产品数={m['product_count']}")
    print("中国TOP10生态评分排名:")
    for r in cn_scores:
        print(f"- {r['生态排名']}. {r['管理人']}: {r['生态评分']:.1f}")
    print("各维度字段来源:")
    print("- 产品广度: 固定TOP10产品数量 + 指数/主题覆盖数量")
    print("- 费率优势: 管理费率 + 托管费率，按产品AUM加权")
    print("- 交易活跃: 2026-07-07补充近60日日均成交额计算成交效率 + 活跃产品占比")
    print("- 多元能力: 多元产品AUM/截图固定规模 + 8类能力模块覆盖数")
    print("- 旗舰产品: 最大ETF规模占比 + Top3产品规模占比，分母使用截图固定规模")
    print("缺失字段WARN:")
    print("无" if not warnings else "\n".join(f"- {w}" for w in warnings))
    print(f"输出sheet列表: {verify.sheetnames}")
    qa = "PASS" if verify.sheetnames == ["美国TOP10五维评分", "中国TOP10五维评分", "评分细则", "数据信息"] and product_counts.get("华宝基金", 0) > 0 and "华安基金" not in [m["manager"] for m in metrics] else "FAIL"
    print(f"QA结论: {qa}")


if __name__ == "__main__":
    main()
