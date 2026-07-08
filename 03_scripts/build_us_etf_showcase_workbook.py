from __future__ import annotations

from copy import copy
from pathlib import Path
from statistics import median

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
    / "美国ETF案例数据收集_富国启示.xlsx"
)
OUTPUT_PATH = (
    Path.home()
    / "Desktop"
    / "Fullgoal"
    / "ICI课题"
    / "课题更正"
    / "美国ETF案例数据收集_富国启示_展示整理版.xlsx"
)

NEW_SHEETS = [
    "00_README",
    "01_Issuer_Overview_Clean",
    "02_Product_Mix_Clean",
    "03_Representative_ETFs_Clean",
    "04_Official_Sources_Clean",
    "05_P23_Case_Map_Data",
    "06_P24_iShares_OS_Data",
    "07_P25_JPMorgan_Migration_Data",
    "08_P25_Invesco_Helper_Data",
    "09_P26_Fullgoal_Implication",
    "10_Slide_Text_Prompts",
]

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


def as_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def hundred_million_usd_to_usd(value):
    number = as_float(value)
    return None if number is None else number * 100_000_000


def hundred_million_usd_to_billion(value):
    number = as_float(value)
    return None if number is None else number / 10


def million_usd_to_usd(value):
    number = as_float(value)
    return None if number is None else number * 1_000_000


def pct_to_decimal(value):
    number = as_float(value)
    return None if number is None else number / 100


def table_rows(wb, sheet_name, header_row=3):
    ws = wb[sheet_name]
    raw_headers = [cell.value for cell in ws[header_row]]
    header_indexes = [(idx, header) for idx, header in enumerate(raw_headers) if header]
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        item = {header: row[idx] for idx, header in header_indexes}
        if any(value not in (None, "") for value in item.values()):
            rows.append(item)
    return rows


def product_key(row):
    return str(row.get("代码") or row.get("ticker") or "").strip().upper()


def infer_product_role(row):
    issuer = row.get("案例角色")
    ticker = product_key(row)
    fund_name = str(row.get("基金名称") or "")
    asset_class = row.get("资产类别")
    strategy = row.get("Strategy")
    tag = row.get("战略标签")
    if tag:
        return tag
    if issuer == "BlackRock/iShares":
        if ticker in {"IVV", "ITOT", "IJH", "IJR"} or "Core" in fund_name:
            if asset_class == "Fixed Income":
                return "Core债券代表产品"
            if "MSCI" in fund_name or "International" in fund_name or ticker in {"IEFA", "IEMG", "IXUS"}:
                return "国际/海外代表产品"
            return "Core宽基代表产品"
        if asset_class == "Fixed Income":
            return "债券/现金管理代表产品"
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


def write_table(wb, sheet_name, headers, rows):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
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

    for idx, cell in enumerate(ws[1], start=1):
        header = str(cell.value or "")
        max_len = len(header)
        for data_cell in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, values_only=False):
            for c in data_cell:
                if c.value is not None:
                    max_len = max(max_len, len(str(c.value)))
        cap = 40 if any(key in header.lower() for key in ["note", "message", "source", "url", "description", "text", "annotation", "说明", "来源", "文案"]) else 24
        width = min(max(max_len + 2, 10), cap)
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22 if row_idx == 1 else 36

    for col_idx, cell in enumerate(ws[1], start=1):
        header = str(cell.value or "")
        fmt = None
        lower = header.lower()
        if header.endswith("_usd") or "amount" in lower or "volume" in lower or "aum_usd" in lower:
            fmt = '$#,##0'
        elif header.endswith("_billion_usd"):
            fmt = '0.00'
        elif "expense_ratio" in lower or "aum_share" in lower:
            fmt = '0.00%'
        if fmt:
            for body_cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                for c in body_cell:
                    c.number_format = fmt
                    c.alignment = Alignment(horizontal="right", vertical="top")


def build_readme():
    return [
        {
            "item": "文件用途",
            "content": "美国ETF案例与富国启示PPT展示型数据包",
            "data_as_of": SNAPSHOT_DATE,
            "source": "用户PPT需求；原始ETF底稿",
        },
        {
            "item": "数据快照日期",
            "content": SNAPSHOT_DATE,
            "data_as_of": SNAPSHOT_DATE,
            "source": "用户指定",
        },
        {
            "item": "数据来源类型",
            "content": "境外ETF全量数据清洗标准化文件；BlackRock/iShares官网；JPMorgan Asset Management官网；Invesco QQQ官网；富国参考PPT/Wind整理数据",
            "data_as_of": SNAPSHOT_DATE,
            "source": "README、Official_Sources、Fullgoal_Reference",
        },
        {
            "item": "量化数据说明",
            "content": "ETF数量、AUM、费率、成交额等量化数据来自境外ETF全量数据清洗标准化文件；AUM/成交额在本工作簿中同时保留USD与billion USD口径。",
            "data_as_of": SNAPSHOT_DATE,
            "source": "Issuer_Overview、Issuer_Product_Mix、Top_Products",
        },
        {
            "item": "定性信息说明",
            "content": "生态打法、模型组合、主动ETF、入口型大单品等定性信息来自官网来源登记及本次PPT逻辑。",
            "data_as_of": SNAPSHOT_DATE,
            "source": "Official_Sources；用户PPT需求",
        },
        {
            "item": "富国侧数据说明",
            "content": "富国侧当前只用于生态位判断；若后续做旗舰池/孵化池/维护池，需要补充富国重点产品成交额、价差、折溢价、净流入等内部或Wind数据。",
            "data_as_of": SNAPSHOT_DATE,
            "source": "Fullgoal_Reference；用户PPT需求",
        },
    ]


def build_issuer_overview(wb):
    overview = table_rows(wb, "Issuer_Overview")
    products = table_rows(wb, "Top_Products")
    rows = []
    for row in overview:
        issuer = row.get("案例角色")
        if issuer not in ISSUER_META:
            continue
        meta = ISSUER_META[issuer]
        top = [
            p for p in products
            if p.get("案例角色") == issuer and as_float(p.get("AUM排名")) and as_float(p.get("AUM排名")) <= 5
        ]
        top.sort(key=lambda x: as_float(x.get("AUM排名")) or 999)
        flagship = "；".join(f"{p.get('代码')} {p.get('基金名称')}" for p in top)
        rows.append({
            "issuer_display_name": issuer,
            "issuer_raw_name": row.get("CSV发行人名称"),
            "case_role": meta["case_role"],
            "etf_count": row.get("ETF数量"),
            "total_aum_usd": hundred_million_usd_to_usd(row.get("合计AUM_亿美元")),
            "total_aum_billion_usd": hundred_million_usd_to_billion(row.get("合计AUM_亿美元")),
            "avg_expense_ratio": pct_to_decimal(row.get("AUM加权费率_%")),
            "median_expense_ratio": None,
            "total_adv_usd": hundred_million_usd_to_usd(row.get("合计日均成交额_亿美元")),
            "flagship_products": flagship or f"{row.get('最大产品代码')} {row.get('最大产品名称')}",
            "ecosystem_keywords": meta["ecosystem_keywords"],
            "ppt_role": meta["ppt_role"],
            "data_as_of": SNAPSHOT_DATE,
            "source_note": f"{row.get('量化来源')}; median_expense_ratio需全量产品级费率数据补充",
        })
    return rows


def build_product_mix(wb):
    rows = []
    for row in table_rows(wb, "Issuer_Product_Mix"):
        issuer = row.get("案例角色")
        if issuer not in ISSUER_META or row.get("拆分维度") != "Asset Class":
            continue
        rows.append({
            "issuer_display_name": issuer,
            "asset_class_or_category": row.get("类别"),
            "etf_count": row.get("ETF数量"),
            "total_aum_usd": hundred_million_usd_to_usd(row.get("AUM_亿美元")),
            "total_aum_billion_usd": hundred_million_usd_to_billion(row.get("AUM_亿美元")),
            "aum_share": pct_to_decimal(row.get("AUM占比_%")),
            "avg_expense_ratio": None,
            "median_expense_ratio": None,
            "ppt_interpretation": ISSUER_META[issuer]["product_mix_note"],
            "data_as_of": SNAPSHOT_DATE,
        })
    return rows


def build_representative_etfs(wb):
    rep_rows = table_rows(wb, "Representative_ETFs")
    top_rows = table_rows(wb, "Top_Products")
    selected = {}

    def add(row, source_sheet):
        ticker = product_key(row)
        issuer = row.get("案例角色")
        if not ticker or issuer not in ISSUER_META:
            return
        selected.setdefault(ticker, (row, source_sheet))

    for row in rep_rows:
        add(row, "Representative_ETFs")
    limits = {"BlackRock/iShares": 10, "JPMorgan AM": 10, "Invesco": 10}
    for row in top_rows:
        issuer = row.get("案例角色")
        rank = as_float(row.get("AUM排名"))
        if issuer in limits and rank is not None and rank <= limits[issuer]:
            add(row, "Top_Products")

    output = []
    for row, source_sheet in selected.values():
        issuer = row.get("案例角色")
        output.append({
            "issuer_display_name": issuer,
            "ticker": product_key(row),
            "fund_name": row.get("基金名称"),
            "case_product_role": infer_product_role(row),
            "asset_class": row.get("资产类别"),
            "category": row.get("Category"),
            "strategy": row.get("Strategy"),
            "aum_usd": hundred_million_usd_to_usd(row.get("AUM_亿美元")),
            "aum_billion_usd": hundred_million_usd_to_billion(row.get("AUM_亿美元")),
            "expense_ratio": pct_to_decimal(row.get("费率_%")),
            "avg_daily_volume_usd": million_usd_to_usd(row.get("日均成交额_百万美元")),
            "ytd_flow_usd": hundred_million_usd_to_usd(row.get("YTD资金流_亿美元")),
            "one_year_flow_usd": hundred_million_usd_to_usd(row.get("近1年资金流_亿美元")),
            "ppt_usage": ISSUER_META[issuer]["ppt_usage"],
            "data_as_of": SNAPSHOT_DATE,
            "source_note": f"{source_sheet}; {row.get('量化来源') or '用户上传ETF.com清洗数据'}",
        })
    output.sort(key=lambda x: (list(ISSUER_META).index(x["issuer_display_name"]), -(x["aum_usd"] or 0)))
    return output


def build_official_sources(wb):
    rows = []
    seen_urls = set()
    for row in table_rows(wb, "Official_Sources"):
        url = row.get("URL")
        if url:
            seen_urls.add(url)
        rows.append({
            "institution": row.get("机构"),
            "source_type": "官网页面",
            "source_title": row.get("资料类型"),
            "url": url,
            "collected_info": row.get("可支撑的表述"),
            "ppt_claim_supported": row.get("可支撑的表述"),
            "recommended_slide": row.get("建议使用页"),
            "data_as_of_or_access_date": SNAPSHOT_DATE,
        })
    additions = [
        {
            "institution": "BlackRock/iShares",
            "source_type": "官网产品列表",
            "source_title": "iShares ETF Investments List",
            "url": "https://www.ishares.com/us/products/etf-investments",
            "collected_info": "iShares ETF产品列表，可支撑产品矩阵和Core ETF工具箱完整度。",
            "ppt_claim_supported": "iShares具备完整ETF产品矩阵和底仓工具箱。",
            "recommended_slide": "P24 产品模块层",
            "data_as_of_or_access_date": SNAPSHOT_DATE,
        },
        {
            "institution": "JPMorgan AM",
            "source_type": "官网产品页",
            "source_title": "JPMorgan BetaBuilders U.S. Equity ETF",
            "url": "https://am.jpmorgan.com/us/en/asset-management/adv/products/jpmorgan-betabuilders-u-s-equity-etf-etf-shares-46641q399",
            "collected_info": "BetaBuilders代表核心Beta能力封装，适合支撑资产配置能力迁移链。",
            "ppt_claim_supported": "JPMorgan将核心Beta能力封装为BetaBuilders ETF。",
            "recommended_slide": "P25 资产配置能力→BetaBuilders/固收ETF",
            "data_as_of_or_access_date": SNAPSHOT_DATE,
        },
        {
            "institution": "Invesco",
            "source_type": "官网ETF产品线",
            "source_title": "Invesco ETF strategies",
            "url": "https://www.invesco.com/us/en/solutions/invesco-etfs.html",
            "collected_info": "Invesco ETF capabilities page lists fixed income, commodities, Innovation Suite, BulletShares, Smart beta and other ETF capabilities.",
            "ppt_claim_supported": "Invesco除QQQ外仍有ETF产品矩阵，但本PPT仅作为入口型单品辅助样本。",
            "recommended_slide": "P25 辅助卡",
            "data_as_of_or_access_date": SNAPSHOT_DATE,
        },
        {
            "institution": "Invesco",
            "source_type": "官网ETF产品列表",
            "source_title": "Exchange-traded funds (ETFs)",
            "url": "https://www.invesco.com/us/en/financial-products/etfs.html",
            "collected_info": "Invesco ETF产品列表，用于复核QQQ以外的Smart Beta、固定收益、商品等代表产品。",
            "ppt_claim_supported": "Invesco具备多类ETF产品线，QQQ是入口心智最强的代表单品。",
            "recommended_slide": "P25 辅助卡",
            "data_as_of_or_access_date": SNAPSHOT_DATE,
        },
    ]
    for addition in additions:
        if addition["url"] not in seen_urls:
            rows.append(addition)
    return rows


def static_case_map():
    source = "研究判断，基于发行人规模、产品结构、官网生态信息及富国参考PPT"
    return [
        {"entity": "BlackRock/iShares", "role": "龙头平台型样本", "x_ecosystem_maturity": 95, "y_asset_management_capability_migration": 85, "bubble_size": 100, "label": "平台化龙头", "annotation": "ETF生态成熟天花板：产品矩阵、Core ETF、Model Portfolios、投顾渠道、账户沉淀。", "source": source},
        {"entity": "JPMorgan AM", "role": "能力迁移型样本", "x_ecosystem_maturity": 75, "y_asset_management_capability_migration": 90, "bubble_size": 70, "label": "能力迁移型追赶者", "annotation": "传统资管机构将主动管理、收益增强、固收与核心Beta能力迁移到ETF。", "source": source},
        {"entity": "Invesco", "role": "入口型辅助样本", "x_ecosystem_maturity": 70, "y_asset_management_capability_migration": 55, "bubble_size": 75, "label": "入口型大单品样本", "annotation": "QQQ作为超级入口建立客户心智，不作为主案例展开。", "source": source},
        {"entity": "富国基金", "role": "广布局追赶者", "x_ecosystem_maturity": 50, "y_asset_management_capability_migration": 70, "bubble_size": 45, "label": "广布局追赶者", "annotation": "从产品铺设转向能力封装，把指数、投研、量化、渠道和交易能力转为ETF解决方案。", "source": source},
    ]


def static_ishares_os():
    return [
        {"layer_order": 1, "layer_name": "ETF产品模块", "layer_keyword": "Core ETF/股票/债券/国际/因子/行业主题", "description": "提供低成本、可组合、覆盖多资产的ETF底层工具箱。", "blackrock_example": "iShares Core ETFs、股票ETF、债券ETF、国际ETF、因子与行业主题ETF", "fullgoal_reference": "筛选核心宽基和重点主题产品作为客户入口。"},
        {"layer_order": 2, "layer_name": "组合构建模块", "layer_keyword": "Model Portfolios/Target Allocation/风险等级组合", "description": "把ETF从单品列表转为组合模板，降低客户和投顾配置门槛。", "blackrock_example": "BlackRock Model Portfolios、Target Allocation、风险等级组合", "fullgoal_reference": "用底仓+主题+策略组合样例承接渠道话术。"},
        {"layer_order": 3, "layer_name": "投顾工具模块", "layer_keyword": "市场观点/组合工具/再平衡/客户风险画像", "description": "通过观点、工具和再平衡流程提升投顾端使用频率。", "blackrock_example": "Advisor Center 360、组合分析、模型比较、税务和再平衡机会", "fullgoal_reference": "围绕重点ETF输出配置观点、组合工具和风险提示。"},
        {"layer_order": 4, "layer_name": "客户账户模块", "layer_keyword": "券商账户/投顾账户/退休账户/长期组合", "description": "让ETF进入客户真实账户和长期配置流程。", "blackrock_example": "投顾账户、券商账户、退休账户中的标准配置模块", "fullgoal_reference": "通过渠道和交易体验把产品转化为账户中的可持续持仓。"},
        {"layer_order": 5, "layer_name": "生态结果模块", "layer_keyword": "规模沉淀/渠道复用/低成本优势/流动性飞轮", "description": "规模、低成本、流动性和渠道复用形成正反馈。", "blackrock_example": "Core ETF规模沉淀、低成本优势、流动性提升和渠道复用", "fullgoal_reference": "集中资源打造旗舰池，提升成交、价差和客户心智。"},
    ]


def static_jpm_migration():
    return [
        {"migration_chain": "主动管理能力 → 主动ETF → 防守/分散/Alpha", "original_capability": "主动管理能力", "etf_packaging": "主动ETF", "representative_products": "JGRO、JCPB、JPST等主动权益/主动固收ETF", "customer_scenario": "防守、分散、Alpha获取", "ppt_message": "传统投研能力可以通过透明、可交易、低门槛的ETF包装进入客户配置场景。", "source": "Representative_ETFs_Clean；JPMorgan Active ETFs官网"},
        {"migration_chain": "收益管理能力 → Equity Premium Income ETF → 月度收入/降低波动/参与部分上行", "original_capability": "收益管理能力", "etf_packaging": "Equity Premium Income ETF", "representative_products": "JEPI、JEPQ", "customer_scenario": "月度收入、降低波动、参与部分上行", "ppt_message": "收益增强能力被压缩为客户容易理解的月度收入型ETF场景。", "source": "Representative_ETFs_Clean；JPMorgan Equity Premium Income ETFs官网"},
        {"migration_chain": "资产配置能力 → BetaBuilders/固收ETF → 底仓/稳健配置/利率配置", "original_capability": "资产配置能力", "etf_packaging": "BetaBuilders/固收ETF", "representative_products": "BBUS、BBJP、BBEU、JPST、JCPB", "customer_scenario": "底仓、稳健配置、利率配置", "ppt_message": "核心Beta和固收能力可转化为底仓与稳健配置工具。", "source": "Representative_ETFs_Clean；JPMorgan BetaBuilders/ETF Investing官网"},
    ]


def build_invesco_helper(rep_rows):
    wanted = []
    for row in rep_rows:
        if row["issuer_display_name"] == "Invesco" and row["ticker"] in {"QQQ", "QQQM"}:
            wanted.append({
                "ticker": row["ticker"],
                "fund_name": row["fund_name"],
                "product_role": row["case_product_role"],
                "aum_usd": row["aum_usd"],
                "aum_billion_usd": row["aum_billion_usd"],
                "expense_ratio": row["expense_ratio"],
                "index_or_exposure": "Nasdaq-100/创新成长",
                "ppt_message": "QQQ把“纳斯达克100/创新成长”压缩成一个ETF符号，对富国的启示是能力封装之外也需要客户记得住的代表产品。",
                "source": row["source_note"],
            })
    wanted.sort(key=lambda x: -(x["aum_usd"] or 0))
    return wanted


def static_fullgoal_implication():
    return [
        {"fullgoal_capability": "指数能力", "current_basis": "宽基、大盘、科创、创业板等布局", "etf_solution_entry": "底仓入口", "customer_scenario": "A股底仓、长期配置、市场Beta获取", "action_suggestion": "筛选1-2只核心宽基作为底仓入口，强化成交、价差、渠道话术和组合示例", "required_followup_data": "重点产品规模、成交额、买卖价差、折溢价、净流入"},
        {"fullgoal_capability": "投研能力", "current_basis": "新能源、信息技术、港股、高端制造、生物医药等主题布局", "etf_solution_entry": "主题入口", "customer_scenario": "产业趋势表达、阶段性行情捕捉、成长弹性配置", "action_suggestion": "持续输出产业观点、配置逻辑和风险提示", "required_followup_data": "主题产品规模、成交、净流入、同类排名、投研内容供给"},
        {"fullgoal_capability": "量化/策略能力", "current_basis": "当前策略ETF关注度较低但可作为中长期差异化方向", "etf_solution_entry": "风格入口", "customer_scenario": "红利、低波、质量、成长、现金流等风格配置", "action_suggestion": "将策略能力转化为客户可理解的收益风险特征", "required_followup_data": "策略ETF产品池、因子暴露、历史表现、回撤、波动、同类排名"},
        {"fullgoal_capability": "渠道与交易能力", "current_basis": "富国已有产品储备，但重点产品心智仍待沉淀", "etf_solution_entry": "使用入口", "customer_scenario": "让渠道会讲、客户敢买、交易顺畅、长期留存", "action_suggestion": "建立旗舰池、孵化池、维护池，集中资源提升重点产品交易体验", "required_followup_data": "做市质量、日均成交额、折溢价、买卖价差、渠道反馈"},
    ]


def static_slide_prompts():
    return [
        {"slide_no": "P23", "slide_title": "美国案例：ETF生态打造的三类路径", "main_message": "美国ETF生态并非只有一种成功路径：龙头靠平台化，传统资管机构靠能力迁移，部分挑战者靠超级单品建立心智。富国更适合重点参考JPMorgan AM的能力迁移路径，并吸收Invesco的大单品心智打法。", "visual_type": "四点战略地图", "left_or_top_text": "三类路径对比：平台化龙头、能力迁移型追赶者、入口型大单品样本。", "chart_data_sheet": "05_P23_Case_Map_Data", "chart_instruction": "横轴生态成熟度，纵轴资管能力迁移程度，气泡大小表示案例权重/规模参考。", "bottom_quote": "富国要看的不是“谁产品最多”，而是谁能让ETF进入客户真实决策流程。", "source_note": "Issuer_Overview_Clean；Official_Sources_Clean；研究判断"},
        {"slide_no": "P24", "slide_title": "BlackRock/iShares：把ETF做成投顾账户的配置操作系统", "main_message": "龙头生态的本质，是把ETF从产品列表变成投顾账户里的标准配置模块。", "visual_type": "五层操作系统架构图", "left_or_top_text": "从ETF产品模块逐层延伸到组合构建、投顾工具、客户账户和生态结果。", "chart_data_sheet": "06_P24_iShares_OS_Data", "chart_instruction": "按layer_order从下到上或从左到右绘制五层架构，每层保留blackrock_example与fullgoal_reference。", "bottom_quote": "iShares的启示：ETF生态的最高形态，不是产品数量领先，而是成为客户资产配置流程中的默认模块。", "source_note": "iShares Core、BlackRock Model Portfolios、Advisor Center官方页面"},
        {"slide_no": "P25", "slide_title": "JPMorgan AM：传统资管机构如何把能力迁移到ETF", "main_message": "JPMorgan不是ETF原生龙头，却把主动管理、收益增强、固定收益和核心Beta能力重新封装成ETF产品和配置场景。", "visual_type": "三条能力迁移链 + Invesco辅助卡片", "left_or_top_text": "左侧三条JPMorgan能力迁移链，右侧放QQQ/QQQM入口心智卡片。", "chart_data_sheet": "07_P25_JPMorgan_Migration_Data + 08_P25_Invesco_Helper_Data", "chart_instruction": "三条链用箭头表现original_capability→etf_packaging→customer_scenario；右侧卡片突出QQQ符号化。", "bottom_quote": "后发ETF管理人不一定靠单一神级产品突围，而是把自身投研、固收、收益增强和渠道能力重新封装成ETF解决方案。", "source_note": "JPMorgan ETF Investing、Active ETFs、Equity Premium Income、BetaBuilders官方页面；Invesco QQQ官网"},
        {"slide_no": "P26", "slide_title": "对富国的启示：从产品铺设转向能力封装", "main_message": "富国当前不是缺ETF产品，而是需要把指数、投研、量化、渠道和交易能力重新封装成客户可理解、可交易、可配置的ETF解决方案。", "visual_type": "富国能力封装飞轮", "left_or_top_text": "四类客户入口：底仓入口、主题入口、风格入口、使用入口。", "chart_data_sheet": "09_P26_Fullgoal_Implication", "chart_instruction": "以富国能力为内圈，以ETF solution entry为外圈，右侧列required_followup_data作为后续提数清单。", "bottom_quote": "富国下一阶段不是继续证明“布局够广”，而是把已有产品储备转化为底仓、主题、策略、交易四类客户入口。", "source_note": "Fullgoal_Reference；用户PPT需求；研究判断"},
    ]


def main():
    wb = load_workbook(INPUT_PATH)
    original_sheets = list(wb.sheetnames)
    for sheet_name in NEW_SHEETS:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    readme_headers = ["item", "content", "data_as_of", "source"]
    overview_headers = ["issuer_display_name", "issuer_raw_name", "case_role", "etf_count", "total_aum_usd", "total_aum_billion_usd", "avg_expense_ratio", "median_expense_ratio", "total_adv_usd", "flagship_products", "ecosystem_keywords", "ppt_role", "data_as_of", "source_note"]
    mix_headers = ["issuer_display_name", "asset_class_or_category", "etf_count", "total_aum_usd", "total_aum_billion_usd", "aum_share", "avg_expense_ratio", "median_expense_ratio", "ppt_interpretation", "data_as_of"]
    rep_headers = ["issuer_display_name", "ticker", "fund_name", "case_product_role", "asset_class", "category", "strategy", "aum_usd", "aum_billion_usd", "expense_ratio", "avg_daily_volume_usd", "ytd_flow_usd", "one_year_flow_usd", "ppt_usage", "data_as_of", "source_note"]
    source_headers = ["institution", "source_type", "source_title", "url", "collected_info", "ppt_claim_supported", "recommended_slide", "data_as_of_or_access_date"]
    case_headers = ["entity", "role", "x_ecosystem_maturity", "y_asset_management_capability_migration", "bubble_size", "label", "annotation", "source"]
    os_headers = ["layer_order", "layer_name", "layer_keyword", "description", "blackrock_example", "fullgoal_reference"]
    migration_headers = ["migration_chain", "original_capability", "etf_packaging", "representative_products", "customer_scenario", "ppt_message", "source"]
    invesco_headers = ["ticker", "fund_name", "product_role", "aum_usd", "aum_billion_usd", "expense_ratio", "index_or_exposure", "ppt_message", "source"]
    fullgoal_headers = ["fullgoal_capability", "current_basis", "etf_solution_entry", "customer_scenario", "action_suggestion", "required_followup_data"]
    slide_headers = ["slide_no", "slide_title", "main_message", "visual_type", "left_or_top_text", "chart_data_sheet", "chart_instruction", "bottom_quote", "source_note"]

    overview_rows = build_issuer_overview(wb)
    rep_clean_rows = build_representative_etfs(wb)

    write_table(wb, "00_README", readme_headers, build_readme())
    write_table(wb, "01_Issuer_Overview_Clean", overview_headers, overview_rows)
    write_table(wb, "02_Product_Mix_Clean", mix_headers, build_product_mix(wb))
    write_table(wb, "03_Representative_ETFs_Clean", rep_headers, rep_clean_rows)
    write_table(wb, "04_Official_Sources_Clean", source_headers, build_official_sources(wb))
    write_table(wb, "05_P23_Case_Map_Data", case_headers, static_case_map())
    write_table(wb, "06_P24_iShares_OS_Data", os_headers, static_ishares_os())
    write_table(wb, "07_P25_JPMorgan_Migration_Data", migration_headers, static_jpm_migration())
    write_table(wb, "08_P25_Invesco_Helper_Data", invesco_headers, build_invesco_helper(rep_clean_rows))
    write_table(wb, "09_P26_Fullgoal_Implication", fullgoal_headers, static_fullgoal_implication())
    write_table(wb, "10_Slide_Text_Prompts", slide_headers, static_slide_prompts())

    wb.save(OUTPUT_PATH)

    print(f"输出文件: {OUTPUT_PATH}")
    print("\n原始sheet保留:")
    for sheet_name in original_sheets:
        print(f"- {sheet_name}")

    print("\n新增sheet行列数（数据行数，不含表头）:")
    for sheet_name in NEW_SHEETS:
        ws = wb[sheet_name]
        print(f"- {sheet_name}: {max(ws.max_row - 1, 0)} 行 x {ws.max_column} 列")

    print("\n三家发行人汇总数据:")
    for row in overview_rows:
        print(
            f"- {row['issuer_display_name']}: ETF数量={row['etf_count']}, "
            f"AUM={row['total_aum_billion_usd']:.2f} billion USD, "
            f"日均成交额={row['total_adv_usd'] / 1_000_000_000:.2f} billion USD, "
            f"AUM加权费率={row['avg_expense_ratio']:.2%}"
        )

    print("\n仍需后续人工补充/复核字段:")
    gaps = [
        "01_Issuer_Overview_Clean.median_expense_ratio：底稿未提供三家全量产品级费率中位数。",
        "02_Product_Mix_Clean.avg_expense_ratio / median_expense_ratio：底稿产品结构聚合表未按资产类别提供费率统计。",
        "09_P26_Fullgoal_Implication.required_followup_data：富国重点产品规模、成交额、买卖价差、折溢价、净流入、做市质量、渠道反馈需后续用Wind/内部数据补齐。",
        "04_Official_Sources_Clean.url：已补充官方页面入口，正式PPT定稿前建议再次人工打开复核页面标题与可访问性。",
    ]
    for gap in gaps:
        print(f"- {gap}")


if __name__ == "__main__":
    main()
