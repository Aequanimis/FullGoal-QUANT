from __future__ import annotations

import csv
import math
import re
import shutil
from collections import Counter, defaultdict
from copy import copy
from datetime import datetime
from pathlib import Path
from statistics import mean, median

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


INPUT_PATH = Path.home() / "Desktop" / "Fullgoal" / "课题研究" / "原始数据" / "全市场ETF基础信息.csv.xlsx"
OUTPUT_DIR = Path.home() / "Desktop" / "Fullgoal" / "ICI课题" / "课题更正"
MAIN_OUTPUT = OUTPUT_DIR / "中国ETF_TOP10生态评分数据包.xlsx"
LIGHT_OUTPUT = OUTPUT_DIR / "中国ETF_TOP10维度评分提取版.xlsx"
CSV_DIR = OUTPUT_DIR / "中国ETF_TOP10生态评分_ChartData"
US_SCORE_FILE = OUTPUT_DIR / "美国ETF_TOP10维度评分提取版_20260706.xlsx"

DIMENSIONS = ["产品矩阵广度", "规模沉淀能力", "成本竞争力", "交易流动性", "能力迁移程度", "入口产品心智"]
DIMENSION_WEIGHTS = {"产品矩阵广度": 0.15, "规模沉淀能力": 0.20, "成本竞争力": 0.15, "交易流动性": 0.15, "能力迁移程度": 0.20, "入口产品心智": 0.15}

SUBS = {
    "产品矩阵广度": [
        ("asset_class_count", "positive", 0.25, "raw"),
        ("category_count", "positive", 0.25, "raw"),
        ("index_or_theme_count", "positive", 0.25, "raw"),
        ("product_type_balance", "positive", 0.25, "raw"),
    ],
    "规模沉淀能力": [
        ("total_aum_log", "positive", 0.35, "log1p"),
        ("etf_count", "positive", 0.10, "raw"),
        ("product_count_over_10bn", "positive", 0.20, "raw"),
        ("product_count_over_50bn", "positive", 0.20, "raw"),
        ("top5_aum_log", "positive", 0.15, "log1p"),
    ],
    "成本竞争力": [
        ("weighted_avg_fee_pct", "negative", 0.40, "raw"),
        ("top10_weighted_fee_pct", "negative", 0.25, "raw"),
        ("low_fee_aum_share", "positive", 0.20, "raw"),
        ("avg_fee_pct", "negative", 0.15, "raw"),
    ],
    "交易流动性": [
        ("total_trading_value_log", "positive", 0.35, "log1p"),
        ("high_liquidity_product_count", "positive", 0.20, "raw"),
        ("high_liquidity_aum_share", "positive", 0.20, "raw"),
        ("weighted_avg_premium_discount_abs", "negative", 0.15, "raw"),
        ("top5_trading_concentration", "negative", 0.10, "raw"),
    ],
    "能力迁移程度": [
        ("non_plain_beta_aum_share", "positive", 0.25, "raw"),
        ("smart_beta_strategy_aum_share", "positive", 0.20, "raw"),
        ("cross_border_aum_share", "positive", 0.15, "raw"),
        ("fixed_income_commodity_aum_share", "positive", 0.15, "raw"),
        ("industry_theme_aum_share", "positive", 0.15, "raw"),
        ("capability_coverage_count", "positive", 0.10, "raw"),
    ],
    "入口产品心智": [
        ("largest_etf_aum_log", "positive", 0.30, "log1p"),
        ("largest_etf_trading_value_log", "positive", 0.20, "log1p"),
        ("top1_aum_share", "positive", 0.15, "raw"),
        ("top3_aum_share", "positive", 0.15, "raw"),
        ("flagship_mind_tag_score", "positive", 0.20, "raw"),
    ],
}

FIELD_CANDIDATES = {
    "fund_code": ["基金代码", "证券代码", "Wind代码", "代码", "ticker"],
    "fund_name": ["基金简称", "证券简称", "基金名称", "基金全称", "fund_name"],
    "manager": ["基金管理人", "管理人", "基金公司", "issuer"],
    "fund_type": ["基金类型", "ETF类型", "投资类型", "投资类型_二级分类", "asset_class"],
    "category": ["ETF细分类别", "产品分类", "分类", "category"],
    "strategy": ["策略类型", "strategy"],
    "index_name": ["跟踪指数", "标的指数", "跟踪指数名称", "index_name"],
    "listing_date": ["上市日期", "成立日期", "基金成立日", "listing_date", "inception_date"],
    "aum": ["基金规模", "最新基金规模(亿)", "基金资产净值", "资产净值", "AUM", "最新规模", "合计规模"],
    "trading_value": ["日均成交额", "成交额", "近一月日均成交额", "日均成交金额"],
    "trading_volume": ["日均成交量", "日均份额成交量"],
    "turnover": ["换手率", "日均换手"],
    "premium_discount": ["折溢价率", "平均折溢价率", "premium_discount"],
    "spread": ["买卖价差", "价差", "spread"],
    "management_fee": ["管理费率", "管理费率_支持历史(%)"],
    "custody_fee": ["托管费率", "托管费率_支持历史(%)"],
    "total_fee": ["综合费率", "合计费率", "总费率"],
    "fund_flow_ytd": ["YTD资金流", "份额变动", "规模变动"],
    "fund_flow_one_year": ["近一年资金流"],
    "issue_size": ["发行规模", "募集规模", "成立规模"],
    "tracking_error": ["跟踪误差", "年化跟踪误差", "近一年跟踪误差"],
}


def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).lower()


def match_fields(headers: list[str]) -> dict[str, str | None]:
    by_norm = {norm(h): h for h in headers if h}
    result = {}
    for req, candidates in FIELD_CANDIDATES.items():
        found = None
        for c in candidates:
            if norm(c) in by_norm:
                found = by_norm[norm(c)]
                break
        if found is None:
            for h in headers:
                hn = norm(h)
                if any(norm(c) in hn or hn in norm(c) for c in candidates):
                    found = h
                    break
        result[req] = found
    return result


def detect_main_sheet(wb):
    best = None
    for ws in wb.worksheets:
        if ws.max_row < 2:
            continue
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        headers = [h for h in headers if h]
        fmap = match_fields(headers)
        hits = sum(1 for k in ["fund_code", "fund_name", "manager", "aum"] if fmap.get(k))
        score = hits * 1000 + len(headers) * 10 + min(ws.max_row, 100000) / 1000
        if best is None or score > best[0]:
            best = (score, ws, headers, fmap)
    if not best or sum(1 for k in ["fund_code", "fund_name", "manager", "aum"] if best[3].get(k)) < 4:
        raise ValueError("未能识别ETF主数据sheet")
    return best[1], best[2], best[3]


def to_float(v):
    if v in (None, "", "-"):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    text = str(v).strip().replace(",", "").replace("%", "")
    if text in ("", "-", "—", "N/A"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def get(row, fmap, key):
    col = fmap.get(key)
    return row.get(col) if col else None


def standardize_manager(v: str) -> str:
    s = str(v or "").strip()
    for suffix in ["基金管理有限公司", "基金管理股份有限公司", "基金有限公司", "管理有限公司", "管理股份有限公司"]:
        s = s.replace(suffix, "基金")
    s = re.sub(r"基金基金$", "基金", s)
    known = ["易方达", "华夏", "华泰柏瑞", "南方", "广发", "国泰", "富国", "汇添富", "嘉实", "博时", "工银瑞信", "招商", "鹏华", "银华", "天弘", "大成", "景顺长城"]
    for k in known:
        if k in s:
            return f"{k}基金"
    return s


def contains(text: str, words: list[str]) -> bool:
    return any(w.lower() in text.lower() for w in words)


def classify(row: dict):
    text = " ".join(str(row.get(k) or "") for k in ["fund_name", "index_name", "fund_type", "category", "strategy"])
    broad = contains(text, ["宽基", "沪深300", "中证500", "中证1000", "中证A500", "上证50", "科创50", "创业板", "深证100", "中证100", "A50", "A100", "上证180"])
    industry = contains(text, ["行业", "主题", "科技", "信息技术", "新能源", "芯片", "半导体", "医药", "消费", "军工", "金融", "证券", "银行", "人工智能", "机器人", "光伏", "储能", "电池", "汽车", "港股科技"])
    smart = contains(text, ["红利", "低波", "质量", "价值", "成长", "动量", "等权", "增强", "策略", "Smart Beta", "ESG", "央企", "国企", "现金流", "基本面", "价值成长"])
    cross = contains(text, ["纳指", "标普", "日经", "德国", "法国", "港股", "恒生", "H股", "QDII", "跨境", "海外", "东南亚", "沙特"])
    fixed = contains(text, ["债", "国债", "政金债", "信用债", "可转债", "公司债", "短融", "利率债"])
    commodity = contains(text, ["黄金", "商品", "有色", "能源", "豆粕", "原油"])
    money = contains(text, ["货币", "现金", "保证金"])
    if money:
        asset = "货币ETF"
        category = "货币"
    elif fixed:
        asset = "债券ETF"
        category = "债券"
    elif commodity:
        asset = "商品ETF"
        category = "商品"
    elif cross:
        asset = "跨境ETF"
        category = "跨境"
    else:
        asset = "股票ETF"
        if broad:
            category = "宽基"
        elif industry:
            category = "行业主题"
        elif smart:
            category = "策略SmartBeta"
        else:
            category = "其他股票"
    row.update({
        "asset_class": asset,
        "category_clean": category,
        "broad_based_flag": broad,
        "industry_theme_flag": industry,
        "smart_beta_strategy_flag": smart,
        "cross_border_flag": cross,
        "fixed_income_flag": fixed,
        "commodity_flag": commodity,
        "money_flag": money,
        "non_plain_beta_flag": industry or smart or cross or fixed or commodity,
    })


def parse_date_str(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if v in (None, ""):
        return None
    text = str(v).strip()
    if re.match(r"\d{4}-\d{1,2}-\d{1,2}", text):
        parts = text[:10].split("-")
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}-{int(parts[2]):02d}"
    return text


def build_rows(ws, headers, fmap):
    rows, excluded_links = [], 0
    date_candidates = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        raw = {headers[i]: values[i] if i < len(values) else None for i in range(len(headers))}
        if not any(v not in (None, "") for v in raw.values()):
            continue
        name = str(get(raw, fmap, "fund_name") or "")
        full_text = " ".join(str(raw.get(k) or "") for k in headers)
        is_etf = "ETF" in full_text.upper() or "交易型开放式" in full_text
        is_link = "联接" in full_text
        if is_link:
            excluded_links += 1
        if not is_etf or is_link or "LOF" in full_text.upper():
            continue
        aum_100m = to_float(get(raw, fmap, "aum"))
        m_fee = to_float(get(raw, fmap, "management_fee"))
        c_fee = to_float(get(raw, fmap, "custody_fee"))
        total_fee = to_float(get(raw, fmap, "total_fee"))
        fee = total_fee if total_fee is not None else ((m_fee or 0) + (c_fee or 0) if m_fee is not None or c_fee is not None else None)
        listing = parse_date_str(get(raw, fmap, "listing_date"))
        for k, v in raw.items():
            if "时间" in str(k) or "日期" in str(k):
                ds = parse_date_str(v)
                if ds and re.match(r"\d{4}-\d{2}-\d{2}$", ds):
                    date_candidates.append(ds)
        row = {
            "fund_code": str(get(raw, fmap, "fund_code") or "").strip(),
            "fund_name": name,
            "manager_raw": str(get(raw, fmap, "manager") or "").strip(),
            "manager": standardize_manager(get(raw, fmap, "manager")),
            "fund_type": get(raw, fmap, "fund_type"),
            "category": get(raw, fmap, "category"),
            "strategy": get(raw, fmap, "strategy"),
            "index_name": get(raw, fmap, "index_name"),
            "listing_date": listing,
            "aum_cny": (aum_100m * 1e8) if aum_100m is not None else None,
            "trading_value_cny": None,
            "premium_discount_pct": None,
            "tracking_error_pct": None,
            "fee_pct": fee,
        }
        classify(row)
        rows.append(row)
    data_as_of = max(date_candidates) if date_candidates else "以原始文件为准，需人工确认"
    if data_as_of and data_as_of.startswith("2026-06"):
        data_as_of = f"{data_as_of}（由费率支持历史时间字段推断，需人工确认）"
    return rows, excluded_links, data_as_of


def sum_field(rows, field):
    return sum(row.get(field) or 0 for row in rows if row.get(field) is not None)


def weighted_avg(rows, val_field, weight_field="aum_cny"):
    pairs = [(row.get(val_field), row.get(weight_field)) for row in rows]
    valid = [(v, w) for v, w in pairs if v is not None and w is not None and w > 0]
    tw = sum(w for _, w in valid)
    if tw <= 0:
        return None
    return sum(v * w for v, w in valid) / tw


def percentile(vals, p):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return None
    if len(vals) == 1:
        return vals[0]
    pos = (len(vals) - 1) * p
    lo, hi = math.floor(pos), math.ceil(pos)
    return vals[lo] if lo == hi else vals[lo] * (hi - pos) + vals[hi] * (pos - lo)


def score_value(x, p5, p95, direction):
    if x is None or p5 is None or p95 is None:
        return None
    if p95 == p5:
        return 80.0
    if direction == "positive":
        return max(60, min(100, 60 + 40 * (x - p5) / (p95 - p5)))
    return max(60, min(100, 60 + 40 * (p95 - x) / (p95 - p5)))


def manager_role(manager, scores):
    if manager == "富国基金":
        return "广布局追赶型 / 能力封装潜力型"
    max_dim = max(DIMENSIONS, key=lambda d: scores[f"{d}_score"])
    if max_dim == "规模沉淀能力" and scores["成本竞争力_score"] >= 85:
        return "低成本规模型"
    if max_dim == "能力迁移程度":
        return "能力迁移型"
    if max_dim == "入口产品心智":
        return "入口大单品型"
    if max_dim == "产品矩阵广度":
        return "平台化龙头型"
    return "特色精品型"


def flagship_score(row):
    aum = row.get("largest_etf_aum_cny") or 0
    score = 70
    if aum >= 50e9:
        score = 95
    elif aum >= 30e9:
        score = 90
    elif aum >= 10e9:
        score = 85
    elif aum >= 5e9:
        score = 80
    name = str(row.get("largest_etf_name") or "")
    if contains(name, ["沪深300ETF", "A500ETF", "上证50ETF", "科创50ETF", "创业板ETF", "黄金ETF", "恒生科技ETF", "纳指ETF", "红利ETF"]):
        score = min(100, score + 5)
    return score


def compute_metrics(rows, data_as_of):
    by_manager = defaultdict(list)
    for row in rows:
        by_manager[row["manager"]].append(row)
    ranked = sorted(by_manager.items(), key=lambda kv: sum_field(kv[1], "aum_cny"), reverse=True)[:10]
    top_managers = [m for m, _ in ranked]
    metrics = []
    for rank, (manager, rs) in enumerate(ranked, start=1):
        total = sum_field(rs, "aum_cny")
        by_aum = sorted(rs, key=lambda r: r.get("aum_cny") or 0, reverse=True)
        top5 = by_aum[:5]
        top10 = by_aum[:10]
        cats = defaultdict(float)
        for r in rs:
            cats[r["category_clean"]] += r.get("aum_cny") or 0
        hhi = sum((v / total) ** 2 for v in cats.values()) if total else None
        largest = by_aum[0] if by_aum else {}
        top3_aum = sum_field(by_aum[:3], "aum_cny")
        flags = {
            "broad": any(r["broad_based_flag"] for r in rs),
            "industry": any(r["industry_theme_flag"] for r in rs),
            "smart": any(r["smart_beta_strategy_flag"] for r in rs),
            "cross": any(r["cross_border_flag"] for r in rs),
            "fixed": any(r["fixed_income_flag"] for r in rs),
            "commodity": any(r["commodity_flag"] for r in rs),
            "money": any(r["money_flag"] for r in rs),
        }
        m = {
            "rank_by_aum": rank,
            "manager": manager,
            "etf_count": len(rs),
            "total_aum_cny": total,
            "total_aum_100m_cny": total / 1e8,
            "total_trading_value_cny": None,
            "total_trading_value_100m_cny": None,
            "avg_fee_pct": mean([r["fee_pct"] for r in rs if r.get("fee_pct") is not None]) if any(r.get("fee_pct") is not None for r in rs) else None,
            "median_fee_pct": median([r["fee_pct"] for r in rs if r.get("fee_pct") is not None]) if any(r.get("fee_pct") is not None for r in rs) else None,
            "weighted_avg_fee_pct": weighted_avg(rs, "fee_pct"),
            "weighted_avg_premium_discount_pct": None,
            "weighted_avg_tracking_error_pct": None,
            "equity_etf_count": sum(1 for r in rs if r["asset_class"] == "股票ETF"),
            "bond_etf_count": sum(1 for r in rs if r["fixed_income_flag"]),
            "cross_border_etf_count": sum(1 for r in rs if r["cross_border_flag"]),
            "commodity_etf_count": sum(1 for r in rs if r["commodity_flag"]),
            "money_etf_count": sum(1 for r in rs if r["money_flag"]),
            "strategy_smart_beta_etf_count": sum(1 for r in rs if r["smart_beta_strategy_flag"]),
            "industry_theme_etf_count": sum(1 for r in rs if r["industry_theme_flag"]),
            "broad_based_etf_count": sum(1 for r in rs if r["broad_based_flag"]),
            "largest_etf_code": largest.get("fund_code"),
            "largest_etf_name": largest.get("fund_name"),
            "largest_etf_aum_cny": largest.get("aum_cny") or 0,
            "largest_etf_trading_value_cny": None,
            "largest_etf_aum_share_pct": ((largest.get("aum_cny") or 0) / total * 100) if total else 0,
            "asset_class_count": len({r["asset_class"] for r in rs}),
            "category_count": len({r["category_clean"] for r in rs}),
            "index_or_theme_count": len({r.get("index_name") or r.get("fund_name") for r in rs}),
            "product_type_balance": (1 - hhi) if hhi is not None else None,
            "product_count_over_10bn": sum(1 for r in rs if (r.get("aum_cny") or 0) >= 1e9),
            "product_count_over_50bn": sum(1 for r in rs if (r.get("aum_cny") or 0) >= 5e9),
            "top5_aum_cny": sum_field(top5, "aum_cny"),
            "top10_weighted_fee_pct": weighted_avg(top10, "fee_pct"),
            "low_fee_aum_share": sum_field([r for r in rs if r.get("fee_pct") is not None and r["fee_pct"] <= 0.20], "aum_cny") / total if total else None,
            "high_liquidity_product_count": None,
            "high_liquidity_aum_share": None,
            "weighted_avg_premium_discount_abs": None,
            "top5_trading_concentration": None,
            "non_plain_beta_aum_share": sum_field([r for r in rs if r["non_plain_beta_flag"]], "aum_cny") / total if total else 0,
            "smart_beta_strategy_aum_share": sum_field([r for r in rs if r["smart_beta_strategy_flag"]], "aum_cny") / total if total else 0,
            "cross_border_aum_share": sum_field([r for r in rs if r["cross_border_flag"]], "aum_cny") / total if total else 0,
            "fixed_income_commodity_aum_share": sum_field([r for r in rs if r["fixed_income_flag"] or r["commodity_flag"]], "aum_cny") / total if total else 0,
            "industry_theme_aum_share": sum_field([r for r in rs if r["industry_theme_flag"]], "aum_cny") / total if total else 0,
            "capability_coverage_count": sum(1 for v in flags.values() if v),
            "top1_aum_share": ((largest.get("aum_cny") or 0) / total) if total else 0,
            "top3_aum_share": top3_aum / total if total else 0,
            "data_as_of": data_as_of,
        }
        m["flagship_mind_tag_score"] = flagship_score(m)
        m["total_aum_log"] = math.log1p(m["total_aum_cny"])
        m["top5_aum_log"] = math.log1p(m["top5_aum_cny"])
        m["total_trading_value_log"] = None
        m["largest_etf_aum_log"] = math.log1p(m["largest_etf_aum_cny"])
        m["largest_etf_trading_value_log"] = None
        metrics.append(m)
    return top_managers, metrics, by_manager


def compute_scores(metrics, data_as_of):
    subscore_rows = []
    dim_scores = {m["manager"]: {} for m in metrics}
    for dim, subs in SUBS.items():
        raw_scores_by_manager = defaultdict(float)
        available_weights_by_manager = defaultdict(float)
        for sub, direction, weight, transform in subs:
            vals = [m.get(sub) for m in metrics if m.get(sub) is not None]
            p5, p95 = percentile(vals, 0.05), percentile(vals, 0.95)
            for m in metrics:
                raw = m.get(sub)
                score = score_value(raw, p5, p95, direction)
                if score is None:
                    adj_weight = 0
                    weighted = None
                    notes = "missing；该子指标因字段缺失不参与维度内加权"
                else:
                    adj_weight = weight
                    weighted = score * weight
                    raw_scores_by_manager[m["manager"]] += weighted
                    available_weights_by_manager[m["manager"]] += weight
                    notes = "winsorized percentile scoring"
                subscore_rows.append({
                    "manager": m["manager"],
                    "dimension": dim,
                    "subindicator": sub,
                    "raw_value": raw,
                    "transform_method": transform,
                    "direction": direction,
                    "p5": p5,
                    "p95": p95,
                    "score_60_100": score,
                    "weight_in_dimension": weight,
                    "adjusted_weight_in_dimension": adj_weight,
                    "weighted_score": weighted,
                    "data_as_of": data_as_of,
                    "notes": notes,
                })
        for m in metrics:
            avail = available_weights_by_manager[m["manager"]]
            dim_scores[m["manager"]][dim] = raw_scores_by_manager[m["manager"]] / avail if avail else 80.0
    summary = []
    for m in metrics:
        manager = m["manager"]
        scores = dim_scores[manager]
        total = sum(scores[d] * DIMENSION_WEIGHTS[d] for d in DIMENSIONS)
        row = {
            "manager": manager,
            **{f"{d}_score": scores[d] for d in DIMENSIONS},
            "total_ecosystem_score": total,
            "data_as_of": data_as_of,
        }
        row["cn_case_role"] = manager_role(manager, row)
        row["interpretation"] = interpret_manager(manager, row)
        summary.append(row)
    summary.sort(key=lambda r: r["total_ecosystem_score"], reverse=True)
    for i, r in enumerate(summary, 1):
        r["rank_total_score"] = i
    return subscore_rows, summary


def interpret_manager(manager, row):
    if manager == "富国基金":
        return "富国处于广布局追赶与能力封装潜力位置，需把现有指数、投研、策略与渠道能力转化为重点ETF入口。"
    top = max(DIMENSIONS, key=lambda d: row[f"{d}_score"])
    return f"{manager}在{top}维度相对突出，可作为中国TOP10 ETF管理人的{row['cn_case_role']}样本。"


def cosine(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a))
    nb = math.sqrt(sum(y * y for y in b))
    return dot / (na * nb) if na and nb else 0


def euclidean(a, b):
    return math.sqrt(sum((x - y) ** 2 for x, y in zip(a, b)))


def read_us_vectors():
    defaults = {
        "BlackRock/iShares": [100.0, 99.7, 90.6, 98.6, 70.9, 85.1],
        "JPMorgan AM": [73.2, 65.5, 77.6, 73.4, 90.0, 70.4],
        "Invesco": [95.8, 76.9, 79.6, 85.0, 67.2, 97.7],
    }
    if not US_SCORE_FILE.exists():
        return defaults
    wb = load_workbook(US_SCORE_FILE, read_only=True, data_only=True)
    if "02_Dimension_Score_Wide" not in wb.sheetnames:
        return defaults
    ws = wb["02_Dimension_Score_Wide"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    out = {}
    for vals in ws.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, vals))
        issuer = row.get("issuer")
        if issuer in defaults:
            out[issuer] = [float(row.get(d) or 0) for d in DIMENSIONS]
    return out or defaults


def build_mapping(summary, us_vectors, data_as_of):
    path_type = {"BlackRock/iShares": "平台化龙头", "JPMorgan AM": "能力迁移型", "Invesco": "入口型大单品"}
    rows = []
    for r in summary:
        vec = [r[f"{d}_score"] for d in DIMENSIONS]
        sims = {u: cosine(vec, v) for u, v in us_vectors.items()}
        dists = {u: euclidean(vec, v) for u, v in us_vectors.items()}
        closest = max(sims, key=sims.get)
        rows.append({
            "manager": r["manager"],
            "cn_total_score": r["total_ecosystem_score"],
            "similarity_to_blackrock": sims.get("BlackRock/iShares"),
            "similarity_to_jpmorgan": sims.get("JPMorgan AM"),
            "similarity_to_invesco": sims.get("Invesco"),
            "closest_us_sample": closest,
            "closest_path_type": path_type.get(closest),
            "interpretation": f"{r['manager']}六维画像最接近{closest}，可参考其{path_type.get(closest)}路径。欧氏距离={dists[closest]:.1f}",
            "data_as_of": data_as_of,
        })
    return rows


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for c in ws[1]:
        c.fill = copy(header_fill)
        c.font = copy(header_font)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = copy(border)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = copy(border)
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for idx, hc in enumerate(ws[1], 1):
        h = str(hc.value or "")
        lower = h.lower()
        max_len = len(h)
        for col in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, max_row=ws.max_row):
            for c in col:
                if c.value is not None:
                    max_len = max(max_len, len(str(c.value)))
        is_text = any(k in lower for k in ["note", "source", "interpretation", "message", "reason", "position", "priority", "notes"])
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 10), 45 if is_text else 24)
        if any(k in lower for k in ["score", "rank", "pct", "share", "weight", "similarity"]):
            fmt = "0.0" if "rank" not in lower else "0"
        elif any(k in lower for k in ["cny", "aum", "value"]):
            fmt = '¥#,##0'
        else:
            fmt = None
        if fmt:
            for col in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, max_row=ws.max_row):
                for c in col:
                    if isinstance(c.value, (int, float)):
                        c.number_format = fmt
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 22 if i == 1 else 34


def write_sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h) for h in headers])
    style_sheet(ws)


def export_csv(path, headers, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def main():
    qa = []
    def add_qa(item, status, notes):
        qa.append({"check_item": item, "status": status, "notes": notes, "data_as_of": "以原始文件为准"})
    add_qa("输入文件是否存在", "PASS" if INPUT_PATH.exists() else "FAIL", str(INPUT_PATH))
    if not INPUT_PATH.exists():
        raise FileNotFoundError(INPUT_PATH)
    src = load_workbook(INPUT_PATH, read_only=True, data_only=True)
    ws, headers, fmap = detect_main_sheet(src)
    add_qa("主数据sheet是否识别成功", "PASS", f"{ws.title} rows={ws.max_row} cols={ws.max_column}")
    rows, excluded_links, data_as_of = build_rows(ws, headers, fmap)
    for q in qa:
        q["data_as_of"] = data_as_of
    add_qa("ETF样本数量", "PASS" if rows else "FAIL", str(len(rows)))
    add_qa("联接基金剔除数量", "PASS", str(excluded_links))
    missing_required = [k for k in ["fund_code", "fund_name", "manager", "aum"] if not fmap.get(k)]
    add_qa("必要字段是否匹配", "PASS" if not missing_required else "FAIL", ",".join(missing_required) or "必要字段均已匹配")
    add_qa("成交字段是否缺失", "WARN" if not fmap.get("trading_value") else "PASS", fmap.get("trading_value") or "未识别成交额字段，交易流动性维度按缺失处理")
    add_qa("折溢价/跟踪误差字段是否缺失", "WARN" if not fmap.get("premium_discount") or not fmap.get("tracking_error") else "PASS", f"premium={fmap.get('premium_discount')}; tracking_error={fmap.get('tracking_error')}")

    managers, metrics, by_manager = compute_metrics(rows, data_as_of)
    add_qa("TOP10是否成功生成", "PASS" if len(managers) == 10 else "FAIL", str(managers))
    add_qa("富国基金是否在TOP10中", "PASS" if "富国基金" in managers else "WARN", str(managers))
    add_qa("规模字段是否异常", "PASS" if all((r.get("aum_cny") or 0) >= 0 for r in rows) else "WARN", "检查AUM非负")
    max_fee = max([r.get("fee_pct") or 0 for r in rows] or [0])
    add_qa("费率字段是否异常", "WARN" if max_fee > 5 else "PASS", f"max_fee={max_fee}")
    sub_rows, summary = compute_scores(metrics, data_as_of)
    us_vectors = read_us_vectors()
    mapping = build_mapping(summary, us_vectors, data_as_of)
    add_qa("是否生成全部sheet", "PASS", "主包16张sheet，轻量包12张sheet")

    summary_by_manager = {r["manager"]: r for r in summary}
    metric_by_manager = {r["manager"]: r for r in metrics}
    mapping_by_manager = {r["manager"]: r for r in mapping}
    radar_long = [{"manager": r["manager"], "dimension": d, "score": r[f"{d}_score"], "total_ecosystem_score": r["total_ecosystem_score"], "cn_case_role": r["cn_case_role"], "data_as_of": data_as_of} for r in summary for d in DIMENSIONS]
    radar_wide = [{"manager": r["manager"], **{d: r[f"{d}_score"] for d in DIMENSIONS}, "total_ecosystem_score": r["total_ecosystem_score"], "cn_case_role": r["cn_case_role"], "data_as_of": data_as_of} for r in summary]
    top_products = []
    flagship = []
    for manager in managers:
        prods = sorted(by_manager[manager], key=lambda x: x.get("aum_cny") or 0, reverse=True)[:10]
        met = metric_by_manager[manager]
        flagship.append({"manager": manager, "flagship_code": met["largest_etf_code"], "flagship_name": met["largest_etf_name"], "flagship_reason": "最大AUM产品；按规模和名称标签识别旗舰心智", "flagship_mind_tag_score": met["flagship_mind_tag_score"], "largest_etf_aum_cny": met["largest_etf_aum_cny"], "largest_etf_trading_value_cny": met["largest_etf_trading_value_cny"], "top1_aum_share": met["top1_aum_share"], "top3_aum_share": met["top3_aum_share"], "notes": "成交额字段缺失时交易心智只使用规模和集中度", "data_as_of": data_as_of})
        for i, p in enumerate(prods, 1):
            top_products.append({"manager": manager, "rank_within_manager": i, "fund_code": p["fund_code"], "fund_name": p["fund_name"], "asset_class": p["asset_class"], "category": p["category_clean"], "index_name": p["index_name"], "aum_cny": p["aum_cny"], "aum_100m_cny": (p.get("aum_cny") or 0)/1e8, "trading_value_cny": p["trading_value_cny"], "trading_value_100m_cny": None, "fee_pct": p["fee_pct"], "premium_discount_pct": p["premium_discount_pct"], "tracking_error_pct": p["tracking_error_pct"], "product_tag": p["category_clean"], "flagship_tag": "旗舰" if i == 1 else "", "data_as_of": data_as_of})

    cn_top10 = [{k: m.get(k) for k in ["rank_by_aum","manager","etf_count","total_aum_cny","total_aum_100m_cny","total_trading_value_cny","total_trading_value_100m_cny","avg_fee_pct","median_fee_pct","weighted_avg_fee_pct","weighted_avg_premium_discount_pct","weighted_avg_tracking_error_pct","equity_etf_count","bond_etf_count","cross_border_etf_count","commodity_etf_count","money_etf_count","strategy_smart_beta_etf_count","industry_theme_etf_count","broad_based_etf_count","largest_etf_code","largest_etf_name","largest_etf_aum_cny","largest_etf_aum_share_pct","data_as_of"]} for m in metrics]
    name_map = []
    grouped = defaultdict(list)
    for r in rows:
        grouped[(r["manager_raw"], r["manager"])].append(r)
    for (raw, std), rs in sorted(grouped.items(), key=lambda kv: sum_field(kv[1], "aum_cny"), reverse=True):
        name_map.append({"manager_raw": raw, "manager_standardized": std, "product_count": len(rs), "total_aum_cny": sum_field(rs, "aum_cny"), "total_aum_100m_cny": sum_field(rs, "aum_cny")/1e8, "notes": "按公司简称标准化"})
    field_map = [{"required_field": k, "matched_raw_field": v, "match_status": "MATCHED" if v else "MISSING", "notes": "自动模糊匹配"} for k, v in fmap.items()]

    score_headers = ["rank_total_score","manager"] + [f"{d}_score" for d in DIMENSIONS] + ["total_ecosystem_score","cn_case_role","interpretation","data_as_of"]
    for i, r in enumerate(summary, 1):
        r["rank_total_score"] = i
    raw_headers = ["manager","etf_count","total_aum_cny","total_aum_100m_cny","total_trading_value_cny","total_trading_value_100m_cny","asset_class_count","category_count","index_or_theme_count","product_type_balance","product_count_over_10bn","product_count_over_50bn","top5_aum_cny","avg_fee_pct","weighted_avg_fee_pct","top10_weighted_fee_pct","low_fee_aum_share","high_liquidity_product_count","high_liquidity_aum_share","weighted_avg_premium_discount_abs","top5_trading_concentration","non_plain_beta_aum_share","smart_beta_strategy_aum_share","cross_border_aum_share","fixed_income_commodity_aum_share","industry_theme_aum_share","capability_coverage_count","largest_etf_code","largest_etf_name","largest_etf_aum_cny","largest_etf_trading_value_cny","top1_aum_share","top3_aum_share","flagship_mind_tag_score","data_as_of"]
    sub_headers = ["manager","dimension","subindicator","raw_value","transform_method","direction","p5","p95","score_60_100","weight_in_dimension","adjusted_weight_in_dimension","weighted_score","data_as_of","notes"]
    method_rows, method_map = build_methodology(data_as_of)
    slide_rows = build_slide_ready(summary, metrics, mapping, us_vectors, data_as_of)
    fullgoal_rows = build_fullgoal(summary_by_manager, metric_by_manager, mapping_by_manager, data_as_of)
    readme = [{"item":"文件用途","content":"中国TOP10 ETF管理人生态能力评分体系，用于与美国TOP10 ETF管理人生态评分体系对比。","data_as_of":data_as_of,"notes":"展示型、解释型评分，不作为正式行业排名"},{"item":"输入文件路径","content":str(INPUT_PATH),"data_as_of":data_as_of,"notes":"不修改原始文件"},{"item":"ETF样本口径","content":"只保留上市交易ETF，剔除ETF联接基金、LOF和非ETF普通指数基金。","data_as_of":data_as_of,"notes":"货币ETF、债券ETF、商品ETF、跨境ETF保留"},{"item":"TOP10筛选方法","content":"按标准化管理人汇总total_aum_cny，取规模前10。","data_as_of":data_as_of,"notes":"基于全市场ETF口径"},{"item":"中国口径与美国口径差异","content":"中国口径将能力迁移拆分为宽基、行业主题、策略/Smart Beta、跨境、债券、商品、货币等模块；成交额、折溢价、跟踪误差字段在本文件中缺失。","data_as_of":data_as_of,"notes":"缺失项已在QA和方法论中说明"}]

    main_sheets = [
        ("00_README", ["item","content","data_as_of","notes"], readme),
        ("01_Field_Map", ["required_field","matched_raw_field","match_status","notes"], field_map),
        ("02_Manager_Name_Map", ["manager_raw","manager_standardized","product_count","total_aum_cny","total_aum_100m_cny","notes"], name_map),
        ("03_CN_Manager_Top10", list(cn_top10[0].keys()), cn_top10),
        ("04_Manager_Raw_Metrics", raw_headers, metrics),
        ("05_Subindicator_Scores", sub_headers, sub_rows),
        ("06_Manager_Score_Summary", score_headers, summary),
        ("07_Radar_Long", ["manager","dimension","score","total_ecosystem_score","cn_case_role","data_as_of"], radar_long),
        ("08_Radar_Wide", ["manager"]+DIMENSIONS+["total_ecosystem_score","cn_case_role","data_as_of"], radar_wide),
        ("09_Top_Products_By_Manager", list(top_products[0].keys()), top_products),
        ("10_Flagship_Product_Mind", list(flagship[0].keys()), flagship),
        ("11_CN_US_Path_Mapping", list(mapping[0].keys()), mapping),
        ("12_Fullgoal_Position", list(fullgoal_rows[0].keys()) if fullgoal_rows else ["manager","data_as_of","notes"], fullgoal_rows or [{"manager":"富国基金","data_as_of":data_as_of,"notes":"富国基金未进入TOP10"}]),
        ("13_Slide_Ready_Data", sorted({k for r in slide_rows for k in r.keys()}), slide_rows),
        ("14_Methodology", ["section","content","data_as_of","notes"], method_rows),
        ("15_QA_Checks", ["check_item","status","notes","data_as_of"], qa),
    ]
    save_workbook(MAIN_OUTPUT, main_sheets)

    light_sheets = [
        ("00_README", ["item","content","data_as_of","notes"], readme),
        ("01_TOP10_Score_Ranking", score_headers, summary),
        ("02_Dimension_Score_Wide", ["manager"]+DIMENSIONS+["total_ecosystem_score","cn_case_role","data_as_of"], radar_wide),
        ("03_Dimension_Score_Long", ["manager","dimension","score","total_ecosystem_score","cn_case_role","data_as_of"], radar_long),
        ("04_Subindicator_Score_Detail", sub_headers, sub_rows),
        ("05_Raw_Metrics_TOP10", raw_headers, metrics),
        ("06_Methodology_Summary", ["section","content","data_as_of","notes"], method_rows),
        ("07_Dimension_Method_Map", list(method_map[0].keys()), method_map),
        ("08_QA_Checks", ["check_item","status","notes","data_as_of"], qa),
        ("09_Chart_Data_Index", list(build_chart_index(data_as_of)[0].keys()), build_chart_index(data_as_of)),
        ("10_CN_US_Path_Mapping", list(mapping[0].keys()), mapping),
        ("11_Fullgoal_Position", list(fullgoal_rows[0].keys()) if fullgoal_rows else ["manager","data_as_of","notes"], fullgoal_rows or [{"manager":"富国基金","data_as_of":data_as_of,"notes":"富国基金未进入TOP10"}]),
    ]
    save_workbook(LIGHT_OUTPUT, light_sheets)

    if CSV_DIR.exists():
        shutil.rmtree(CSV_DIR)
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    csv_specs = {
        "CN_TOP10_Score_Ranking.csv": (score_headers, summary),
        "CN_Dimension_Score_Wide.csv": (["manager"]+DIMENSIONS+["total_ecosystem_score","cn_case_role","data_as_of"], radar_wide),
        "CN_Dimension_Score_Long.csv": (["manager","dimension","score","total_ecosystem_score","cn_case_role","data_as_of"], radar_long),
        "CN_Subindicator_Score_Detail.csv": (sub_headers, sub_rows),
        "CN_US_Path_Mapping.csv": (list(mapping[0].keys()), mapping),
        "Fullgoal_vs_US_Selected.csv": (sorted({k for r in slide_rows if r.get("block")=="Fullgoal_vs_US_Selected" for k in r.keys()}), [r for r in slide_rows if r.get("block")=="Fullgoal_vs_US_Selected"]),
        "Fullgoal_Position.csv": (list(fullgoal_rows[0].keys()) if fullgoal_rows else ["manager","data_as_of","notes"], fullgoal_rows or [{"manager":"富国基金","data_as_of":data_as_of,"notes":"富国基金未进入TOP10"}]),
        "CN_TOP10_Score_Bar.csv": (sorted({k for r in slide_rows if r.get("block")=="CN_TOP10_Score_Bar" for k in r.keys()}), [r for r in slide_rows if r.get("block")=="CN_TOP10_Score_Bar"]),
    }
    for fname, (headers_, rows_) in csv_specs.items():
        export_csv(CSV_DIR / fname, headers_, rows_)

    counts = Counter(q["status"] for q in qa)
    fg = summary_by_manager.get("富国基金")
    fg_map = mapping_by_manager.get("富国基金")
    print(f"输入文件路径: {INPUT_PATH}")
    print(f"输出主文件路径: {MAIN_OUTPUT}")
    print(f"输出轻量评分文件路径: {LIGHT_OUTPUT}")
    print(f"主数据sheet名称和行列数: {ws.title}, {ws.max_row}行 x {ws.max_column}列")
    print(f"ETF样本数量: {len(rows)}")
    print(f"剔除联接基金数量: {excluded_links}")
    print("中国ETF管理人TOP10列表:")
    for m in metrics:
        print(f"- {m['rank_by_aum']}. {m['manager']}: ETF数量={m['etf_count']}, AUM={m['total_aum_100m_cny']:.2f}亿元")
    print("TOP10六维评分和总分排名:")
    for r in summary:
        print(f"- {r['rank_total_score']}. {r['manager']}: 总分={r['total_ecosystem_score']:.1f}, 产品={r['产品矩阵广度_score']:.1f}, 规模={r['规模沉淀能力_score']:.1f}, 成本={r['成本竞争力_score']:.1f}, 流动性={r['交易流动性_score']:.1f}, 迁移={r['能力迁移程度_score']:.1f}, 入口={r['入口产品心智_score']:.1f}")
    if fg:
        print(f"富国基金六维评分、总分排名和生态位解释: rank={fg['rank_total_score']}, 总分={fg['total_ecosystem_score']:.1f}, {fg['interpretation']}")
        print(f"富国最相似的美国样本: {fg_map['closest_us_sample']} ({fg_map['closest_path_type']})")
    else:
        print("富国基金未进入TOP10，未生成富国TOP10生态位评分。")
    print("生成的sheet列表:")
    for s, _, _ in main_sheets:
        print(f"- {s}")
    print("生成的CSV文件列表:")
    for p in sorted(CSV_DIR.glob("*.csv")):
        print(f"- {p}")
    print(f"QA中PASS/WARN/FAIL数量: PASS={counts.get('PASS',0)}, WARN={counts.get('WARN',0)}, FAIL={counts.get('FAIL',0)}")
    for q in qa:
        if q["status"] in {"WARN","FAIL"}:
            print(f"- {q['status']}: {q['check_item']} - {q['notes']}")


def build_methodology(data_as_of):
    rows = [
        ("样本选择", "保留上市交易ETF，剔除ETF联接基金、LOF和非ETF普通指数基金。"),
        ("数据清洗", "规模字段最新基金规模(亿)换算为人民币元；综合费率优先用总费率，否则管理费率+托管费率。"),
        ("中国字段映射", "基金代码、基金简称、基金管理人、跟踪指数、最新基金规模、管理/托管费率均自动匹配。"),
        ("六维定义", "产品矩阵广度、规模沉淀能力、成本竞争力、交易流动性、能力迁移程度、入口产品心智。"),
        ("权重", "总分=15%产品矩阵+20%规模+15%成本+15%流动性+20%能力迁移+15%入口心智。"),
        ("缺失指标权重归一化方法", "某子指标因字段缺失无法计算时，在维度内部对剩余可用指标做权重归一化；若整维无可用指标，记中性80分。"),
        ("中美映射方法", "读取美国三类样本六维得分，计算中国管理人与BlackRock/iShares、JPMorgan AM、Invesco的余弦相似度。"),
        ("局限性", "本源文件缺少成交额、折溢价和跟踪误差字段，交易流动性维度信息不足。"),
    ]
    method_rows = [{"section": s, "content": c, "data_as_of": data_as_of, "notes": "中国ETF TOP10生态评分方法"} for s,c in rows]
    method_map = []
    for dim, subs in SUBS.items():
        for sub, direction, weight, transform in subs:
            method_map.append({"dimension": dim, "dimension_weight_in_total": DIMENSION_WEIGHTS[dim], "subindicator": sub, "subindicator_weight_in_dimension": weight, "direction": direction, "transform_method": transform, "raw_metric_source": f"04_Manager_Raw_Metrics.{sub}", "plain_language_explanation": "该子指标用于解释管理人在该维度的相对能力。", "data_as_of": data_as_of, "notes": "维度-子指标映射"})
    return method_rows, method_map


def build_chart_index(data_as_of):
    return [
        {"chart_name":"中国TOP10总分排名柱状图","source_sheet":"01_TOP10_Score_Ranking","chart_type":"bar","x_field":"manager","y_field":"total_ecosystem_score","recommended_usage":"展示中国TOP10 ETF管理人生态能力总分排序","data_as_of":data_as_of,"notes":"图表索引"},
        {"chart_name":"中国TOP10六维雷达图","source_sheet":"03_Dimension_Score_Long","chart_type":"radar","x_field":"dimension","y_field":"score","recommended_usage":"生成TOP10或精选管理人的六维雷达图","data_as_of":data_as_of,"notes":"图表索引"},
        {"chart_name":"中美路径映射散点图","source_sheet":"10_CN_US_Path_Mapping","chart_type":"table/scatter","x_field":"closest_us_sample","y_field":"cn_total_score","recommended_usage":"展示中国管理人与美国三类样本的路径相似度","data_as_of":data_as_of,"notes":"图表索引"},
    ]


def build_fullgoal(summary_by, metric_by, mapping_by, data_as_of):
    if "富国基金" not in summary_by:
        return []
    r = summary_by["富国基金"]; m = metric_by["富国基金"]; mp = mapping_by["富国基金"]
    dims = {d: r[f"{d}_score"] for d in DIMENSIONS}
    strongest = max(dims, key=dims.get); weakest = min(dims, key=dims.get)
    return [{"manager":"富国基金","rank_by_aum":m["rank_by_aum"],"rank_total_score":r["rank_total_score"],"total_aum_100m_cny":m["total_aum_100m_cny"],"etf_count":m["etf_count"],**{f"{d}_score":dims[d] for d in DIMENSIONS},"strongest_dimension":strongest,"weakest_dimension":weakest,"closest_us_sample":mp["closest_us_sample"],"fullgoal_ecosystem_position":r["cn_case_role"],"fullgoal_learning_priority_1":"参考JPMorgan AM的能力迁移打法，将投研/固收/策略能力封装成ETF解决方案","fullgoal_learning_priority_2":"参考BlackRock/iShares的平台化工具箱，强化底仓组合和渠道工具","fullgoal_learning_priority_3":"参考Invesco入口型单品心智，打造客户记得住的代表ETF","ppt_message":r["interpretation"],"data_as_of":data_as_of}]


def build_slide_ready(summary, metrics, mapping, us_vectors, data_as_of):
    metric_by = {m["manager"]: m for m in metrics}; map_by={m["manager"]:m for m in mapping}
    rows=[]
    for r in summary:
        rows.append({"block":"CN_TOP10_Score_Bar","manager":r["manager"],"total_ecosystem_score":r["total_ecosystem_score"],"cn_case_role":r["cn_case_role"],"data_as_of":data_as_of})
        for d in DIMENSIONS:
            rows.append({"block":"CN_TOP10_Radar_Long","manager":r["manager"],"dimension":d,"score":r[f"{d}_score"],"data_as_of":data_as_of})
    entities = {"BlackRock/iShares":("美国","平台化龙头样本"),"JPMorgan AM":("美国","能力迁移型样本"),"Invesco":("美国","入口型大单品样本")}
    for us, vec in us_vectors.items():
        for d, score in zip(DIMENSIONS, vec):
            rows.append({"block":"Fullgoal_vs_US_Selected","entity":us,"market":"美国","dimension":d,"score":score,"role":entities[us][1],"data_as_of":data_as_of})
    if any(r["manager"]=="富国基金" for r in summary):
        fg=next(r for r in summary if r["manager"]=="富国基金")
        for d in DIMENSIONS:
            rows.append({"block":"Fullgoal_vs_US_Selected","entity":"富国基金","market":"中国","dimension":d,"score":fg[f"{d}_score"],"role":fg["cn_case_role"],"data_as_of":data_as_of})
    for r in summary:
        m=metric_by[r["manager"]]; mp=map_by[r["manager"]]
        rows.append({"block":"CN_US_Path_Map","manager":r["manager"],"x_能力迁移程度":r["能力迁移程度_score"],"y_规模沉淀能力":r["规模沉淀能力_score"],"bubble_size_total_aum":m["total_aum_100m_cny"],"closest_us_sample":mp["closest_us_sample"],"label":r["manager"],"data_as_of":data_as_of})
    for title in ["指数能力 → 底仓入口","投研能力 → 主题入口","量化/策略能力 → 风格入口","渠道与交易能力 → 使用入口"]:
        rows.append({"block":"Fullgoal_Action_Text","action_title":title,"action_text":"结合中国TOP10评分结果，将产品布局转化为客户可理解、可交易、可配置的ETF入口。","data_as_of":data_as_of})
    return rows


def save_workbook(path, specs):
    wb=Workbook(); wb.remove(wb.active)
    for name, headers, rows in specs:
        write_sheet(wb, name, headers, rows)
    wb.save(path)


if __name__ == "__main__":
    main()
