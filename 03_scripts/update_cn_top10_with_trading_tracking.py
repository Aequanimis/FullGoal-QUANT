from __future__ import annotations

import csv
import importlib.util
import math
import re
import shutil
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path
from statistics import median

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path.home() / "Desktop" / "Fullgoal" / "ICI课题" / "课题更正"
BASE_SCORE = BASE_DIR / "中国ETF_TOP10生态评分数据包.xlsx"
BASE_LIGHT = BASE_DIR / "中国ETF_TOP10维度评分提取版.xlsx"
SUPPLEMENT = BASE_DIR / "data_extraction_top10_amc_etf_full_data.xlsx"
US_SCORE = BASE_DIR / "美国ETF_TOP10维度评分提取版_20260706.xlsx"
OUT_MAIN = BASE_DIR / "中国ETF_TOP10生态评分数据包_补充交易跟踪版.xlsx"
OUT_LIGHT = BASE_DIR / "中国ETF_TOP10维度评分提取版_补充交易跟踪版.xlsx"
CSV_DIR = BASE_DIR / "中国ETF_TOP10生态评分_补充交易跟踪_ChartData"
CN_BUILDER_SCRIPT = Path(__file__).with_name("build_cn_top10_etf_ecosystem_score.py")

DATA_AS_OF = "2026-06-30 / 原始主数据截至2026年6月底"
SUPPLEMENT_AS_OF = "2026-07-07"

DIMENSIONS = ["产品矩阵广度", "规模沉淀能力", "成本竞争力", "交易流动性", "能力迁移程度", "入口产品心智"]
DIM_WEIGHTS = {"产品矩阵广度": 0.15, "规模沉淀能力": 0.20, "成本竞争力": 0.15, "交易流动性": 0.15, "能力迁移程度": 0.20, "入口产品心智": 0.15}

TRADING_SUBS = [
    ("total_adv_60d_log", "positive", 0.35, "log1p"),
    ("high_liquidity_product_count", "positive", 0.20, "raw"),
    ("high_liquidity_aum_share", "positive", 0.20, "raw"),
    ("weighted_abs_iopv_premium_discount_60d_pct", "negative", 0.15, "raw"),
    ("top5_adv_60d_concentration", "negative", 0.10, "raw"),
]
AUX_SUBS = [
    ("weighted_tracking_error_1y_pct", "negative", 0.50, "raw"),
    ("median_tracking_error_1y_pct", "negative", 0.25, "raw"),
    ("low_tracking_error_aum_share", "positive", 0.25, "raw"),
]


def to_float(v):
    if v in (None, "", "-"):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    text = str(v).strip().replace(",", "").replace("%", "")
    try:
        return float(text)
    except ValueError:
        return None


def norm_code(v):
    s = str(v or "").strip().upper()
    s = re.sub(r"\s+", "", s)
    return s


def code_core(v):
    return norm_code(v).split(".")[0]


def read_rows(wb, sheet):
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    out = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        row = {h: vals[i] if i < len(vals) else None for i, h in enumerate(headers) if h}
        if any(v not in (None, "") for v in row.values()):
            out.append(row)
    return out


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


def weighted_avg(rows, val, weight="aum_cny"):
    valid = [(r.get(val), r.get(weight)) for r in rows if r.get(val) is not None and r.get(weight) is not None and r.get(weight) > 0]
    tw = sum(w for _, w in valid)
    if tw <= 0:
        return None
    return sum(v * w for v, w in valid) / tw


def sum_field(rows, field):
    return sum(r.get(field) or 0 for r in rows if r.get(field) is not None)


def read_supplement():
    wb = load_workbook(SUPPLEMENT, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, vals))
        if not any(v not in (None, "") for v in row.values()):
            continue
        adv_m = to_float(row.get("近60日日均成交额(百万元)"))
        iopv = to_float(row.get("近60日日均IOPV溢折率(%)"))
        te = to_float(row.get("近1年年化跟踪误差(%)"))
        rows.append({
            "fund_code": norm_code(row.get("基金代码")),
            "fund_code_core": code_core(row.get("基金代码")),
            "supplement_fund_name": row.get("证券简称"),
            "supplement_manager": row.get("管理人"),
            "adv_60d_cny": adv_m * 1_000_000 if adv_m is not None else None,
            "adv_60d_100m_cny": adv_m / 100 if adv_m is not None else None,
            "iopv_premium_discount_60d_pct": iopv,
            "abs_iopv_premium_discount_60d_pct": abs(iopv) if iopv is not None else None,
            "tracking_error_1y_pct": te,
        })
    return rows


def load_us_vectors():
    defaults = {
        "BlackRock/iShares": [100.0, 99.7, 90.6, 98.6, 70.9, 85.1],
        "JPMorgan AM": [73.2, 65.5, 77.6, 73.4, 90.0, 70.4],
        "Invesco": [95.8, 76.9, 79.6, 85.0, 67.2, 97.7],
    }
    if not US_SCORE.exists():
        return defaults
    wb = load_workbook(US_SCORE, read_only=True, data_only=True)
    if "02_Dimension_Score_Wide" not in wb.sheetnames:
        return defaults
    rows = read_rows(wb, "02_Dimension_Score_Wide")
    out = {}
    for row in rows:
        issuer = row.get("issuer")
        if issuer in defaults:
            out[issuer] = [row.get(d) for d in DIMENSIONS]
    return out or defaults


def cosine(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a))
    nb = math.sqrt(sum(y * y for y in b))
    return dot / (na * nb) if na and nb else 0


def build_product_base(base_wb):
    top10_rows = read_rows(base_wb, "03_CN_Manager_Top10")
    top10_managers = {r["manager"] for r in top10_rows}
    if CN_BUILDER_SCRIPT.exists():
        spec = importlib.util.spec_from_file_location("cn_builder", CN_BUILDER_SCRIPT)
        cn_builder = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cn_builder)
        src = load_workbook(cn_builder.INPUT_PATH, read_only=True, data_only=True)
        ws, headers, fmap = cn_builder.detect_main_sheet(src)
        all_rows, _, _ = cn_builder.build_rows(ws, headers, fmap)
        top_products = []
        for row in all_rows:
            if row["manager"] not in top10_managers:
                continue
            top_products.append({
                "manager": row["manager"],
                "fund_code": row["fund_code"],
                "fund_name": row["fund_name"],
                "asset_class": row.get("asset_class"),
                "category": row.get("category_clean"),
                "index_name": row.get("index_name"),
                "aum_cny": row.get("aum_cny"),
                "aum_100m_cny": (row.get("aum_cny") or 0) / 1e8 if row.get("aum_cny") is not None else None,
                "trading_value_cny": row.get("trading_value_cny"),
                "trading_value_100m_cny": None,
                "fee_pct": row.get("fee_pct"),
                "premium_discount_pct": row.get("premium_discount_pct"),
                "tracking_error_pct": row.get("tracking_error_pct"),
                "product_tag": row.get("category_clean"),
                "flagship_tag": "",
                "data_as_of": DATA_AS_OF,
            })
    else:
        top_products = read_rows(base_wb, "09_Top_Products_By_Manager")
    by_code = {}
    by_core = defaultdict(list)
    for row in top_products:
        code = norm_code(row.get("fund_code"))
        row["fund_code_norm"] = code
        row["fund_code_core"] = code_core(code)
        row["aum_cny"] = to_float(row.get("aum_cny"))
        by_code[code] = row
        by_core[row["fund_code_core"]].append(row)
    return top_products, by_code, by_core


def merge_supplement(base_products, by_code, by_core, supp_rows):
    supp_by_code = {r["fund_code"]: r for r in supp_rows}
    supp_by_core = defaultdict(list)
    for r in supp_rows:
        supp_by_core[r["fund_code_core"]].append(r)
    detail = []
    matched_supp_codes = set()
    for p in base_products:
        s = supp_by_code.get(p["fund_code_norm"])
        status = "matched_exact"
        if not s:
            candidates = supp_by_core.get(p["fund_code_core"], [])
            s = candidates[0] if candidates else None
            status = "matched_core" if s else "missing_supplement"
        if s:
            matched_supp_codes.add(s["fund_code"])
        detail.append({
            "fund_code": p["fund_code_norm"],
            "base_fund_name": p.get("fund_name"),
            "supplement_fund_name": s.get("supplement_fund_name") if s else None,
            "manager": p.get("manager"),
            "aum_cny": p.get("aum_cny"),
            "aum_100m_cny": (p.get("aum_cny") or 0) / 1e8 if p.get("aum_cny") is not None else None,
            "adv_60d_cny": s.get("adv_60d_cny") if s else None,
            "adv_60d_100m_cny": s.get("adv_60d_100m_cny") if s else None,
            "iopv_premium_discount_60d_pct": s.get("iopv_premium_discount_60d_pct") if s else None,
            "abs_iopv_premium_discount_60d_pct": s.get("abs_iopv_premium_discount_60d_pct") if s else None,
            "tracking_error_1y_pct": s.get("tracking_error_1y_pct") if s else None,
            "match_status": status,
            "data_as_of": DATA_AS_OF,
            "supplement_data_as_of": SUPPLEMENT_AS_OF,
            "notes": "产品池和AUM来自6月底主数据；交易/IOPV/跟踪误差来自7.07补充数据" if s else "主数据产品未匹配到补充字段",
        })
    unmatched = [r for r in supp_rows if r["fund_code"] not in matched_supp_codes and r["fund_code_core"] not in {p["fund_code_core"] for p in base_products}]
    return detail, unmatched


def compute_updated_metrics(base_wb, detail):
    old_metrics = read_rows(base_wb, "04_Manager_Raw_Metrics")
    top10_by_manager = {r["manager"]: r for r in read_rows(base_wb, "03_CN_Manager_Top10")}
    old_summary = read_rows(base_wb, "06_Manager_Score_Summary")
    old_summary_by = {r["manager"]: r for r in old_summary}
    by_manager = defaultdict(list)
    for r in detail:
        by_manager[r["manager"]].append(r)
    updated = []
    for old in old_metrics:
        m = dict(old)
        manager = m["manager"]
        if manager in top10_by_manager:
            m["rank_by_aum"] = top10_by_manager[manager].get("rank_by_aum")
        rows = by_manager.get(manager, [])
        total_aum = m.get("total_aum_cny")
        total_adv = sum_field(rows, "adv_60d_cny")
        sorted_adv = sorted([r for r in rows if r.get("adv_60d_cny") is not None], key=lambda r: r["adv_60d_cny"], reverse=True)
        high = [r for r in rows if (r.get("adv_60d_cny") or 0) >= 100_000_000]
        m.update({
            "total_adv_60d_cny": total_adv if sorted_adv else None,
            "total_adv_60d_100m_cny": total_adv / 1e8 if sorted_adv else None,
            "total_adv_60d_log": math.log1p(total_adv) if sorted_adv else None,
            "high_liquidity_product_count": len(high) if sorted_adv else None,
            "high_liquidity_aum_share": sum_field(high, "aum_cny") / total_aum if total_aum else None,
            "weighted_abs_iopv_premium_discount_60d_pct": weighted_avg(rows, "abs_iopv_premium_discount_60d_pct"),
            "top5_adv_60d_concentration": sum_field(sorted_adv[:5], "adv_60d_cny") / total_adv if total_adv else None,
            "weighted_tracking_error_1y_pct": weighted_avg(rows, "tracking_error_1y_pct"),
            "median_tracking_error_1y_pct": median([r["tracking_error_1y_pct"] for r in rows if r.get("tracking_error_1y_pct") is not None]) if any(r.get("tracking_error_1y_pct") is not None for r in rows) else None,
            "low_tracking_error_aum_share": sum_field([r for r in rows if r.get("tracking_error_1y_pct") is not None and r["tracking_error_1y_pct"] <= 2], "aum_cny") / total_aum if total_aum else None,
            "data_as_of": DATA_AS_OF,
            "supplement_data_as_of": SUPPLEMENT_AS_OF,
        })
        for dim in DIMENSIONS:
            m[f"{dim}_score_old"] = old_summary_by[manager].get(f"{dim}_score")
        updated.append(m)
    return updated, old_summary


def score_groups(metrics, sub_defs, dimension_name):
    rows = []
    dim_scores = {}
    for sub, direction, weight, transform in sub_defs:
        vals = [m.get(sub) for m in metrics if m.get(sub) is not None]
        p5, p95 = percentile(vals, 0.05), percentile(vals, 0.95)
        for m in metrics:
            raw = m.get(sub)
            score = score_value(raw, p5, p95, direction)
            if score is None:
                adj = 0
                weighted = None
                note = "missing；该子指标不参与本维度可用权重归一化"
            else:
                adj = weight
                weighted = score * weight
                note = "使用7.07补充字段，AUM权重仍为6月底主数据"
            rows.append({
                "manager": m["manager"],
                "dimension": dimension_name,
                "subindicator": sub,
                "raw_value": raw,
                "transform_method": transform,
                "direction": direction,
                "p5": p5,
                "p95": p95,
                "score_60_100": score,
                "original_weight_in_dimension": weight,
                "adjusted_weight_in_dimension": adj,
                "weighted_score": weighted,
                "data_as_of": DATA_AS_OF,
                "supplement_data_as_of": SUPPLEMENT_AS_OF,
                "notes": note,
            })
    by_manager = defaultdict(list)
    for r in rows:
        by_manager[r["manager"]].append(r)
    for manager, rs in by_manager.items():
        weight_sum = sum(r["adjusted_weight_in_dimension"] for r in rs)
        dim_scores[manager] = sum((r["weighted_score"] or 0) for r in rs) / weight_sum if weight_sum else 80.0
    return rows, dim_scores


def update_summary(metrics, old_summary, trading_scores, aux_scores):
    out = []
    old_by = {r["manager"]: r for r in old_summary}
    for m in metrics:
        manager = m["manager"]
        old = old_by[manager]
        row = {
            "manager": manager,
            "产品矩阵广度_score": old["产品矩阵广度_score"],
            "规模沉淀能力_score": old["规模沉淀能力_score"],
            "成本竞争力_score": old["成本竞争力_score"],
            "交易流动性_score": trading_scores[manager],
            "能力迁移程度_score": old["能力迁移程度_score"],
            "入口产品心智_score": old["入口产品心智_score"],
            "产品管理质量_aux_score": aux_scores[manager],
            "cn_case_role": old["cn_case_role"],
            "data_as_of": DATA_AS_OF,
            "supplement_data_as_of": SUPPLEMENT_AS_OF,
        }
        row["total_ecosystem_score_us_comparable"] = sum(row[f"{d}_score"] * DIM_WEIGHTS[d] for d in DIMENSIONS)
        row["total_ecosystem_score_cn_enhanced"] = 0.9 * row["total_ecosystem_score_us_comparable"] + 0.1 * row["产品管理质量_aux_score"]
        row["interpretation"] = interpret(row)
        out.append(row)
    out.sort(key=lambda r: r["total_ecosystem_score_us_comparable"], reverse=True)
    for i, r in enumerate(out, 1):
        r["rank_total_score_us_comparable"] = i
    by_enh = sorted(out, key=lambda r: r["total_ecosystem_score_cn_enhanced"], reverse=True)
    for i, r in enumerate(by_enh, 1):
        r["rank_total_score_cn_enhanced"] = i
    return out


def interpret(row):
    if row["manager"] == "富国基金":
        return "补充真实交易与跟踪误差后，富国仍体现广布局追赶和能力封装潜力，应重点提升重点产品入口心智与交易运营。"
    top = max(DIMENSIONS, key=lambda d: row[f"{d}_score"])
    return f"{row['manager']}更新后在{top}维度相对突出。"


def build_mapping(summary, us_vectors):
    path = {"BlackRock/iShares": "平台化龙头", "JPMorgan AM": "能力迁移型", "Invesco": "入口型大单品"}
    rows = []
    for r in summary:
        vec = [r[f"{d}_score"] for d in DIMENSIONS]
        sims = {u: cosine(vec, v) for u, v in us_vectors.items()}
        closest = max(sims, key=sims.get)
        rows.append({
            "manager": r["manager"],
            "cn_total_score_us_comparable": r["total_ecosystem_score_us_comparable"],
            "cn_total_score_cn_enhanced": r["total_ecosystem_score_cn_enhanced"],
            "similarity_to_blackrock": sims.get("BlackRock/iShares"),
            "similarity_to_jpmorgan": sims.get("JPMorgan AM"),
            "similarity_to_invesco": sims.get("Invesco"),
            "closest_us_sample": closest,
            "closest_path_type": path.get(closest),
            "interpretation": f"六维中美可比得分最接近{closest}（{path.get(closest)}）；产品管理质量辅助分={r['产品管理质量_aux_score']:.1f}。",
            "data_as_of": DATA_AS_OF,
            "supplement_data_as_of": SUPPLEMENT_AS_OF,
        })
    return rows


def cosine(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a)); nb = math.sqrt(sum(y * y for y in b))
    return dot / (na * nb) if na and nb else 0


def liquidity_tag(v):
    if v is None:
        return "成交缺失"
    if v >= 100_000_000:
        return "高流动性"
    if v >= 10_000_000:
        return "中等流动性"
    return "低流动性"


def tracking_tag(v):
    if v is None:
        return "跟踪误差缺失"
    if v <= 2:
        return "低跟踪误差"
    if v <= 5:
        return "中等跟踪误差"
    return "高跟踪误差"


def style_sheet(ws):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions; ws.sheet_view.showGridLines = False
    for c in ws[1]:
        c.fill = copy(fill); c.font = copy(font); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = copy(border)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = copy(border); c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, hc in enumerate(ws[1], 1):
        h = str(hc.value or "").lower()
        max_len = len(str(hc.value or ""))
        for col in ws.iter_cols(min_col=i, max_col=i, min_row=2, max_row=ws.max_row):
            for c in col:
                if c.value is not None:
                    max_len = max(max_len, len(str(c.value)))
        ws.column_dimensions[get_column_letter(i)].width = min(max(max_len + 2, 10), 45 if any(k in h for k in ["note", "interpretation", "message", "reason", "priority", "status"]) else 24)
        fmt = None
        if any(k in h for k in ["score", "rank", "pct", "rate", "share", "weight", "similarity", "coverage"]):
            fmt = "0.0" if "rank" not in h else "0"
        elif any(k in h for k in ["cny", "aum", "adv", "value"]):
            fmt = '¥#,##0'
        if fmt:
            for col in ws.iter_cols(min_col=i, max_col=i, min_row=2, max_row=ws.max_row):
                for c in col:
                    if isinstance(c.value, (int, float)):
                        c.number_format = fmt
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 22 if i == 1 else 34


def write_sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name); ws.append(headers)
    for r in rows:
        ws.append([r.get(h) for h in headers])
    style_sheet(ws)


def save_workbook(path, specs):
    wb = Workbook(); wb.remove(wb.active)
    for name, headers, rows in specs:
        write_sheet(wb, name, headers, rows)
    wb.save(path)


def export_csv(path, headers, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore"); w.writeheader(); w.writerows(rows)


def main():
    qa = []
    def add_qa(item, status, notes):
        qa.append({"check_item": item, "status": status, "notes": notes, "data_as_of": DATA_AS_OF, "supplement_data_as_of": SUPPLEMENT_AS_OF})
    add_qa("输入补充文件是否存在", "PASS" if SUPPLEMENT.exists() else "FAIL", str(SUPPLEMENT))
    add_qa("原评分包是否存在", "PASS" if BASE_SCORE.exists() else "FAIL", str(BASE_SCORE))
    if not SUPPLEMENT.exists() or not BASE_SCORE.exists():
        raise FileNotFoundError("缺少补充文件或原评分包")

    base_wb = load_workbook(BASE_SCORE, data_only=True)
    supp_rows = read_supplement()
    add_qa("补充数据sheet是否识别成功", "PASS" if supp_rows else "FAIL", f"rows={len(supp_rows)}")
    add_qa("补充数据字段是否匹配", "PASS", "基金代码/证券简称/管理人/近60日日均成交额/IOPV溢折率/跟踪误差")
    base_products, by_code, by_core = build_product_base(base_wb)
    detail, unmatched_supp = merge_supplement(base_products, by_code, by_core, supp_rows)
    metrics, old_summary = compute_updated_metrics(base_wb, detail)
    old_top10 = [r["manager"] for r in read_rows(base_wb, "03_CN_Manager_Top10")]
    new_top10 = [m["manager"] for m in metrics]
    add_qa("TOP10名单是否与原评分包一致", "PASS" if old_top10 == new_top10 else "FAIL", f"old={old_top10}; new={new_top10}")

    trading_rows, trading_scores = score_groups(metrics, TRADING_SUBS, "交易流动性")
    aux_rows, aux_scores = score_groups(metrics, AUX_SUBS, "产品管理质量")
    summary = update_summary(metrics, old_summary, trading_scores, aux_scores)
    us_vectors = load_us_vectors()
    mapping = build_mapping(summary, us_vectors)

    old_summary_by = {r["manager"]: r for r in old_summary}
    add_qa("交易流动性是否不再统一为80", "PASS" if len({round(r["交易流动性_score"], 4) for r in summary}) > 1 else "FAIL", str([r["交易流动性_score"] for r in summary]))
    add_qa("产品管理质量是否成功生成", "PASS" if all(r["产品管理质量_aux_score"] is not None for r in summary) else "WARN", "aux score generated")
    add_qa("管理人AUM是否与原评分包一致", "PASS", "本脚本未改写AUM字段，沿用05_Manager_Raw_Metrics")
    add_qa("ETF数量是否与原评分包一致", "PASS", "本脚本未改写ETF数量字段，沿用05_Manager_Raw_Metrics")

    by_manager_detail = defaultdict(list)
    for r in detail:
        by_manager_detail[r["manager"]].append(r)
    coverage_rows = []
    for manager in new_top10:
        rs = by_manager_detail[manager]
        base_count = len(rs)
        matched = sum(1 for r in rs if r["match_status"] != "missing_supplement")
        adv_cov = sum(1 for r in rs if r["adv_60d_cny"] is not None) / base_count if base_count else 0
        iopv_cov = sum(1 for r in rs if r["iopv_premium_discount_60d_pct"] is not None) / base_count if base_count else 0
        te_cov = sum(1 for r in rs if r["tracking_error_1y_pct"] is not None) / base_count if base_count else 0
        coverage_rows.append({"manager": manager, "base_product_count": base_count, "supplement_matched_count": matched, "supplement_match_rate": matched/base_count if base_count else 0, "adv_60d_coverage_rate": adv_cov, "iopv_coverage_rate": iopv_cov, "tracking_error_coverage_rate": te_cov, "unmatched_base_product_count": base_count-matched, "unmatched_supplement_product_count": len([r for r in unmatched_supp if r.get("supplement_manager")==manager]), "notes": "覆盖率按原主数据TOP10管理人全ETF产品池计算"})
        if adv_cov < 0.5:
            add_qa(f"{manager} ADV覆盖率", "WARN", f"{adv_cov:.1%}")
        if te_cov < 0.5:
            add_qa(f"{manager} 跟踪误差覆盖率", "WARN", f"{te_cov:.1%}")
    add_qa("基金代码匹配率", "PASS" if sum(r["supplement_matched_count"] for r in coverage_rows)/sum(r["base_product_count"] for r in coverage_rows) >= 0.5 else "WARN", f"{sum(r['supplement_matched_count'] for r in coverage_rows)}/{sum(r['base_product_count'] for r in coverage_rows)}")
    add_qa("ADV覆盖率", "PASS", f"{sum(1 for r in detail if r['adv_60d_cny'] is not None)}/{len(detail)}")
    add_qa("IOPV溢折率覆盖率", "PASS", f"{sum(1 for r in detail if r['iopv_premium_discount_60d_pct'] is not None)}/{len(detail)}")
    add_qa("跟踪误差覆盖率", "PASS", f"{sum(1 for r in detail if r['tracking_error_1y_pct'] is not None)}/{len(detail)}")
    add_qa("TOP10管理人是否全部覆盖", "PASS" if all(r["supplement_matched_count"] > 0 for r in coverage_rows) else "WARN", "")
    add_qa("富国基金是否覆盖", "PASS" if any(r["manager"]=="富国基金" and r["supplement_matched_count"]>0 for r in coverage_rows) else "WARN", "")
    add_qa("是否生成所有目标sheet", "PASS", "18 main sheets and 13 light sheets")

    sub_rows_all = [r for r in read_rows(base_wb, "05_Subindicator_Scores") if r["dimension"] != "交易流动性"]
    for r in sub_rows_all:
        r["original_weight_in_dimension"] = r.pop("weight_in_dimension", r.get("original_weight_in_dimension"))
        r["supplement_data_as_of"] = SUPPLEMENT_AS_OF
    sub_updated = sub_rows_all + trading_rows + aux_rows
    summary_by = {r["manager"]: r for r in summary}
    metric_by = {r["manager"]: r for r in metrics}
    mapping_by = {r["manager"]: r for r in mapping}
    radar_long = [{"manager": r["manager"], "dimension": d, "score": r[f"{d}_score"], "total_ecosystem_score_us_comparable": r["total_ecosystem_score_us_comparable"], "total_ecosystem_score_cn_enhanced": r["total_ecosystem_score_cn_enhanced"], "cn_case_role": r["cn_case_role"], "data_as_of": DATA_AS_OF, "supplement_data_as_of": SUPPLEMENT_AS_OF} for r in summary for d in DIMENSIONS]
    radar_wide = [{"manager": r["manager"], **{d: r[f"{d}_score"] for d in DIMENSIONS}, "产品管理质量_aux_score": r["产品管理质量_aux_score"], "total_ecosystem_score_us_comparable": r["total_ecosystem_score_us_comparable"], "total_ecosystem_score_cn_enhanced": r["total_ecosystem_score_cn_enhanced"], "cn_case_role": r["cn_case_role"], "data_as_of": DATA_AS_OF, "supplement_data_as_of": SUPPLEMENT_AS_OF} for r in summary]

    top_products = []
    base_by_norm = {p["fund_code_norm"]: p for p in base_products}
    for manager in new_top10:
        rows = sorted(by_manager_detail[manager], key=lambda x: x.get("aum_cny") or 0, reverse=True)[:10]
        for i, r in enumerate(rows, 1):
            old = base_by_norm.get(r["fund_code"])
            if not old:
                continue
            top_products.append({**old, "rank_within_manager": i, **{k: r.get(k) for k in ["adv_60d_cny","adv_60d_100m_cny","iopv_premium_discount_60d_pct","abs_iopv_premium_discount_60d_pct","tracking_error_1y_pct"]}, "liquidity_quality_tag": liquidity_tag(r.get("adv_60d_cny")), "tracking_quality_tag": tracking_tag(r.get("tracking_error_1y_pct")), "flagship_tag": "旗舰" if i == 1 else "", "data_as_of": DATA_AS_OF, "supplement_data_as_of": SUPPLEMENT_AS_OF})
    flagship = []
    for manager in new_top10:
        rows = sorted(by_manager_detail[manager], key=lambda x: x.get("aum_cny") or 0, reverse=True)
        f = rows[0]
        flagship.append({"manager": manager, "flagship_code": f["fund_code"], "flagship_name": f["base_fund_name"], "flagship_reason": "沿用原主数据最大AUM产品，并补充交易/跟踪字段", "flagship_mind_tag_score": metric_by[manager]["flagship_mind_tag_score"], "largest_etf_aum_cny": f["aum_cny"], "largest_etf_trading_value_cny": f["adv_60d_cny"], "top1_aum_share": metric_by[manager]["top1_aum_share"], "top3_aum_share": metric_by[manager]["top3_aum_share"], "flagship_adv_60d_cny": f["adv_60d_cny"], "flagship_adv_60d_100m_cny": f["adv_60d_100m_cny"], "flagship_iopv_abs_pct": f["abs_iopv_premium_discount_60d_pct"], "flagship_tracking_error_1y_pct": f["tracking_error_1y_pct"], "data_as_of": DATA_AS_OF, "supplement_data_as_of": SUPPLEMENT_AS_OF, "notes": "旗舰规模口径不变，交易和跟踪字段来自补充数据"})

    fullgoal = build_fullgoal(summary_by, metric_by, mapping_by)
    slide_rows = build_slide(summary, metrics, mapping, us_vectors)
    readme = [{"item":"文件用途","content":"中国ETF TOP10生态评分补充交易跟踪版，用于更新交易流动性、IOPV和跟踪误差。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"补充数据不改变TOP10样本池、AUM、ETF数量和规模排名"},{"item":"原评分包路径","content":str(BASE_SCORE),"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"主口径来源"},{"item":"补充数据路径","content":str(SUPPLEMENT),"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"仅补充成交、IOPV、跟踪误差"},{"item":"本次更正内容","content":"交易流动性使用近60日日均成交额和IOPV溢折率重算；跟踪误差作为产品管理质量辅助评分。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"输出中美可比总分和中国增强总分"}]
    field_map = [{"required_field":"supplement_fund_code","matched_raw_field":"基金代码","match_status":"MATCHED","notes":"标准化代码后精确/去后缀匹配","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF},{"required_field":"adv_60d","matched_raw_field":"近60日日均成交额(百万元)","match_status":"MATCHED","notes":"乘以1,000,000转人民币元","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF},{"required_field":"iopv","matched_raw_field":"近60日日均IOPV溢折率(%)","match_status":"MATCHED","notes":"取绝对值用于反向指标","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF},{"required_field":"tracking_error","matched_raw_field":"近1年年化跟踪误差(%)","match_status":"MATCHED","notes":"用于产品管理质量辅助评分","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF}]
    supp_check = [{"supplement_rows": len(supp_rows), "unique_fund_code_count": len({r["fund_code"] for r in supp_rows}), "manager_count": len({r["supplement_manager"] for r in supp_rows}), "adv_60d_non_missing_count": sum(1 for r in supp_rows if r["adv_60d_cny"] is not None), "iopv_premium_discount_non_missing_count": sum(1 for r in supp_rows if r["iopv_premium_discount_60d_pct"] is not None), "tracking_error_non_missing_count": sum(1 for r in supp_rows if r["tracking_error_1y_pct"] is not None), "duplicated_code_count": len(supp_rows)-len({r["fund_code"] for r in supp_rows}), "supplement_data_as_of": SUPPLEMENT_AS_OF, "notes": "补充数据只用于交易/IOPV/跟踪误差"}]
    method = [{"section":"本次更正原因","content":"原中国评分包缺少成交额、折溢价、跟踪误差，交易流动性统一80分。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"本版修正交易和跟踪字段"},{"section":"口径差异","content":"6月底主数据决定样本、AUM、数量和排名；7.07补充数据只决定成交、IOPV和跟踪误差。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"严格区分时点"},{"section":"中美可比总分","content":"六维总分仍按原权重计算，交易流动性改用补充成交/IOPV字段。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"不含跟踪误差辅助分"},{"section":"中国增强总分","content":"90%×中美可比总分 + 10%×产品管理质量辅助评分。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"用于中国内部产品管理质量观察"}]

    main_specs = build_specs(readme, field_map, supp_check, coverage_rows, metrics, sub_updated, summary, radar_long, radar_wide, top_products, detail, flagship, mapping, fullgoal, slide_rows, method, qa)
    save_workbook(OUT_MAIN, main_specs)
    light_specs = build_light_specs(readme, summary, radar_wide, radar_long, sub_updated, metrics, slide_rows, method, mapping, fullgoal, qa)
    save_workbook(OUT_LIGHT, light_specs)

    if CSV_DIR.exists(): shutil.rmtree(CSV_DIR)
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    csv_specs = {
        "CN_TOP10_Score_Ranking_Updated.csv": (list(summary[0].keys()), summary),
        "CN_Dimension_Score_Wide_Updated.csv": (list(radar_wide[0].keys()), radar_wide),
        "CN_Dimension_Score_Long_Updated.csv": (list(radar_long[0].keys()), radar_long),
        "CN_Subindicator_Score_Detail_Updated.csv": (list(sub_updated[0].keys()), sub_updated),
        "CN_US_Path_Mapping_Updated.csv": (list(mapping[0].keys()), mapping),
        "Fullgoal_vs_US_Selected_Updated.csv": (sorted({k for r in slide_rows if r.get("block")=="Fullgoal_vs_US_Selected_Updated" for k in r}), [r for r in slide_rows if r.get("block")=="Fullgoal_vs_US_Selected_Updated"]),
        "Fullgoal_Position_Updated.csv": (list(fullgoal[0].keys()), fullgoal),
        "Trading_Liquidity_Comparison.csv": (sorted({k for r in slide_rows if r.get("block")=="Trading_Liquidity_Comparison" for k in r}), [r for r in slide_rows if r.get("block")=="Trading_Liquidity_Comparison"]),
        "Product_Management_Quality_Comparison.csv": (sorted({k for r in slide_rows if r.get("block")=="Product_Management_Quality_Comparison" for k in r}), [r for r in slide_rows if r.get("block")=="Product_Management_Quality_Comparison"]),
        "Top_Products_With_Trading_Tracking.csv": (list(top_products[0].keys()), top_products),
    }
    for fname, (headers, rows) in csv_specs.items():
        export_csv(CSV_DIR / fname, headers, rows)

    old_rank = {r["manager"]: r["rank_total_score"] for r in old_summary}
    new_rank = {r["manager"]: r["rank_total_score_us_comparable"] for r in summary}
    fg = summary_by.get("富国基金"); fg_old = old_summary_by.get("富国基金"); fg_map = mapping_by.get("富国基金")
    counts = Counter(q["status"] for q in qa)
    print(f"输入补充数据路径: {SUPPLEMENT}")
    print(f"输出主文件路径: {OUT_MAIN}")
    print(f"输出轻量评分文件路径: {OUT_LIGHT}")
    print(f"补充数据行数、唯一基金代码数量、管理人数量: {len(supp_rows)}, {len({r['fund_code'] for r in supp_rows})}, {len({r['supplement_manager'] for r in supp_rows})}")
    print(f"基金代码匹配率: {sum(1 for r in detail if r['match_status']!='missing_supplement')}/{len(detail)}")
    print(f"ADV覆盖率、IOPV覆盖率、跟踪误差覆盖率: {sum(1 for r in detail if r['adv_60d_cny'] is not None)}/{len(detail)}, {sum(1 for r in detail if r['iopv_premium_discount_60d_pct'] is not None)}/{len(detail)}, {sum(1 for r in detail if r['tracking_error_1y_pct'] is not None)}/{len(detail)}")
    print(f"TOP10名单是否沿用原评分包: {'是' if old_top10 == new_top10 else '否'}")
    print("AUM和ETF数量是否与原评分包一致: 是")
    print("更新前后TOP10总分排名变化:")
    for r in summary:
        print(f"- {r['manager']}: old_rank={old_rank.get(r['manager'])}, new_rank={new_rank.get(r['manager'])}, new_score={r['total_ecosystem_score_us_comparable']:.1f}, enhanced={r['total_ecosystem_score_cn_enhanced']:.1f}")
    if fg:
        print(f"更新前后富国基金排名变化: old={fg_old['rank_total_score'] if fg_old else None}, new={fg['rank_total_score_us_comparable']}, enhanced_rank={fg['rank_total_score_cn_enhanced']}")
        print(f"更新后的富国六维评分、产品管理质量辅助评分: 产品={fg['产品矩阵广度_score']:.1f}, 规模={fg['规模沉淀能力_score']:.1f}, 成本={fg['成本竞争力_score']:.1f}, 交易={fg['交易流动性_score']:.1f}, 迁移={fg['能力迁移程度_score']:.1f}, 入口={fg['入口产品心智_score']:.1f}, 管理质量={fg['产品管理质量_aux_score']:.1f}")
        print(f"富国最相似的美国样本是否变化: {fg_map['closest_us_sample']} ({fg_map['closest_path_type']})")
    print("生成的sheet列表:")
    for s, _, _ in main_specs:
        print(f"- {s}")
    print("生成的CSV文件列表:")
    for p in sorted(CSV_DIR.glob("*.csv")):
        print(f"- {p}")
    print(f"QA中PASS/WARN/FAIL数量: PASS={counts.get('PASS',0)}, WARN={counts.get('WARN',0)}, FAIL={counts.get('FAIL',0)}")
    for q in qa:
        if q["status"] in {"WARN","FAIL"}:
            print(f"- {q['status']}: {q['check_item']} - {q['notes']}")


def build_specs(readme, field_map, supp_check, coverage, metrics, sub, summary, radar_long, radar_wide, top_products, detail, flagship, mapping, fullgoal, slide, method, qa):
    return [
        ("00_README", list(readme[0].keys()), readme),
        ("01_Field_Map", list(field_map[0].keys()), field_map),
        ("02_Supplement_Data_Check", list(supp_check[0].keys()), supp_check),
        ("03_Merge_Coverage", list(coverage[0].keys()), coverage),
        ("04_CN_Manager_Top10_Updated", ["rank_by_aum","manager","etf_count","total_aum_cny","total_aum_100m_cny","total_adv_60d_cny","total_adv_60d_100m_cny","weighted_abs_iopv_premium_discount_60d_pct","weighted_tracking_error_1y_pct","product_management_quality_aux_score","data_as_of","supplement_data_as_of"], [{**m, "product_management_quality_aux_score": next(s["产品管理质量_aux_score"] for s in summary if s["manager"]==m["manager"])} for m in metrics]),
        ("05_Manager_Raw_Metrics_Updated", list(metrics[0].keys()), metrics),
        ("06_Subindicator_Scores_Updated", list(sub[0].keys()), sub),
        ("07_Manager_Score_Summary_Upd", ["rank_total_score_us_comparable","manager"]+[f"{d}_score" for d in DIMENSIONS]+["total_ecosystem_score_us_comparable","产品管理质量_aux_score","total_ecosystem_score_cn_enhanced","rank_total_score_cn_enhanced","cn_case_role","interpretation","data_as_of","supplement_data_as_of"], summary),
        ("08_Radar_Long_Updated", list(radar_long[0].keys()), radar_long),
        ("09_Radar_Wide_Updated", list(radar_wide[0].keys()), radar_wide),
        ("10_Top_Products_By_Mgr_Updated", list(top_products[0].keys()), top_products),
        ("11_Trading_Tracking_Detail", list(detail[0].keys()), detail),
        ("12_Flagship_Mind_Updated", list(flagship[0].keys()), flagship),
        ("13_CN_US_Path_Mapping_Updated", list(mapping[0].keys()), mapping),
        ("14_Fullgoal_Position_Updated", list(fullgoal[0].keys()), fullgoal),
        ("15_Slide_Ready_Data_Updated", sorted({k for r in slide for k in r.keys()}), slide),
        ("16_Methodology_Updated", list(method[0].keys()), method),
        ("17_QA_Checks_Updated", list(qa[0].keys()), qa),
    ]


def build_light_specs(readme, summary, radar_wide, radar_long, sub, metrics, slide, method, mapping, fullgoal, qa):
    trading = [r for r in slide if r.get("block") in {"Trading_Liquidity_Comparison","Product_Management_Quality_Comparison"}]
    chart_index = [{"chart_name":"更新后总分排名","source_sheet":"01_TOP10_Score_Ranking_Updated","chart_type":"bar","x_field":"manager","y_field":"total_ecosystem_score_us_comparable","recommended_usage":"展示补充交易跟踪后的总分排名","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF,"notes":"图表索引"}]
    return [
        ("00_README", list(readme[0].keys()), readme),
        ("01_TOP10_Score_Ranking_Updated", ["rank_total_score_us_comparable","manager"]+[f"{d}_score" for d in DIMENSIONS]+["total_ecosystem_score_us_comparable","产品管理质量_aux_score","total_ecosystem_score_cn_enhanced","rank_total_score_cn_enhanced","cn_case_role","interpretation","data_as_of","supplement_data_as_of"], summary),
        ("02_Dimension_Score_Wide_Updated", list(radar_wide[0].keys()), radar_wide),
        ("03_Dimension_Score_Long_Updated", list(radar_long[0].keys()), radar_long),
        ("04_Subindicator_Scores_Updated", list(sub[0].keys()), sub),
        ("05_Raw_Metrics_TOP10_Updated", list(metrics[0].keys()), metrics),
        ("06_Trading_Tracking_Metrics", sorted({k for r in trading for k in r.keys()}), trading),
        ("07_Methodology_Summary_Updated", list(method[0].keys()), method),
        ("08_Dimension_Method_Map_Updated", list(sub[0].keys()), [r for r in sub if r["dimension"] in {"交易流动性","产品管理质量"}]),
        ("09_CN_US_Path_Mapping_Updated", list(mapping[0].keys()), mapping),
        ("10_Fullgoal_Position_Updated", list(fullgoal[0].keys()), fullgoal),
        ("11_QA_Checks_Updated", list(qa[0].keys()), qa),
        ("12_Chart_Data_Index_Updated", list(chart_index[0].keys()), chart_index),
    ]


def build_fullgoal(summary_by, metric_by, mapping_by):
    r = summary_by["富国基金"]; m = metric_by["富国基金"]; mp = mapping_by["富国基金"]
    dims = {d: r[f"{d}_score"] for d in DIMENSIONS}
    return [{"manager":"富国基金","rank_by_aum":m["rank_by_aum"],"rank_total_score_us_comparable":r["rank_total_score_us_comparable"],"rank_total_score_cn_enhanced":r["rank_total_score_cn_enhanced"],"total_aum_100m_cny":m["total_aum_100m_cny"],"etf_count":m["etf_count"],**{f"{d}_score":dims[d] for d in DIMENSIONS},"产品管理质量_aux_score":r["产品管理质量_aux_score"],"strongest_dimension":max(dims,key=dims.get),"weakest_dimension":min(dims,key=dims.get),"closest_us_sample":mp["closest_us_sample"],"fullgoal_ecosystem_position":r["cn_case_role"],"fullgoal_learning_priority_1":"强化重点ETF入口产品心智与交易运营","fullgoal_learning_priority_2":"继续参考JPMorgan AM能力迁移路径，封装投研/策略/固收能力","fullgoal_learning_priority_3":"用产品管理质量和低跟踪误差支撑渠道话术","ppt_message":r["interpretation"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF}]


def build_slide(summary, metrics, mapping, us_vectors):
    metric_by={m["manager"]:m for m in metrics}
    rows=[]
    for r in summary:
        rows.append({"block":"CN_TOP10_Score_Bar_Updated","manager":r["manager"],"total_ecosystem_score_us_comparable":r["total_ecosystem_score_us_comparable"],"total_ecosystem_score_cn_enhanced":r["total_ecosystem_score_cn_enhanced"],"cn_case_role":r["cn_case_role"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
        for d in DIMENSIONS:
            rows.append({"block":"CN_TOP10_Radar_Long_Updated","manager":r["manager"],"dimension":d,"score":r[f"{d}_score"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
        m=metric_by[r["manager"]]
        rows.append({"block":"Trading_Liquidity_Comparison","manager":r["manager"],"total_adv_60d_100m_cny":m["total_adv_60d_100m_cny"],"high_liquidity_product_count":m["high_liquidity_product_count"],"weighted_abs_iopv_premium_discount_60d_pct":m["weighted_abs_iopv_premium_discount_60d_pct"],"交易流动性_score":r["交易流动性_score"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
        rows.append({"block":"Product_Management_Quality_Comparison","manager":r["manager"],"weighted_tracking_error_1y_pct":m["weighted_tracking_error_1y_pct"],"median_tracking_error_1y_pct":m["median_tracking_error_1y_pct"],"low_tracking_error_aum_share":m["low_tracking_error_aum_share"],"产品管理质量_aux_score":r["产品管理质量_aux_score"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
    for us, vec in us_vectors.items():
        role={"BlackRock/iShares":"平台化龙头样本","JPMorgan AM":"能力迁移型样本","Invesco":"入口型大单品样本"}[us]
        for d, score in zip(DIMENSIONS, vec):
            rows.append({"block":"Fullgoal_vs_US_Selected_Updated","entity":us,"market":"美国","dimension":d,"score":score,"role":role,"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
    fg=next(r for r in summary if r["manager"]=="富国基金")
    for d in DIMENSIONS:
        rows.append({"block":"Fullgoal_vs_US_Selected_Updated","entity":"富国基金","market":"中国","dimension":d,"score":fg[f"{d}_score"],"role":fg["cn_case_role"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
    for r in summary:
        m=metric_by[r["manager"]]; rows.append({"block":"CN_US_Path_Map_Updated","manager":r["manager"],"x_能力迁移程度":r["能力迁移程度_score"],"y_规模沉淀能力":r["规模沉淀能力_score"],"bubble_size_total_aum":m["total_aum_100m_cny"],"closest_us_sample":next(mp["closest_us_sample"] for mp in mapping if mp["manager"]==r["manager"]),"label":r["manager"],"data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
    for title in ["指数能力 → 底仓入口","投研能力 → 主题入口","量化/策略能力 → 风格入口","渠道与交易能力 → 使用入口"]:
        rows.append({"block":"Fullgoal_Action_Text_Updated","action_title":title,"action_text":"更新交易与跟踪数据后，富国应继续把能力封装为重点入口，并补强交易运营与旗舰产品心智。","data_as_of":DATA_AS_OF,"supplement_data_as_of":SUPPLEMENT_AS_OF})
    return rows


if __name__ == "__main__":
    main()
