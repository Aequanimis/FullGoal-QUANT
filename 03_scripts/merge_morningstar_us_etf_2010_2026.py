from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


WORK_DIR = Path(r"C:\Users\lvdon\Desktop\Fullgoal\日常工作\7.8\Morning Star")
FILE1 = WORK_DIR / "美国ETF2010-2026.xlsx"
FILE2 = WORK_DIR / "美国ETF2010-2026-2.xlsx"
OUTPUT_XLSX = WORK_DIR / "美国ETF2010-2026_合并版.xlsx"
OUTPUT_LOG = WORK_DIR / "美国ETF2010-2026_合并日志.txt"

MERGED_SHEET = "美国ETF2010-2026_合并"
SUMMARY_SHEET = "合并说明"
SUMMARY_MARKER = "Summary Statistics"
STAT_MARKERS = {
    "Summary Statistics",
    "Eightieth Percentile",
    "Average",
    "Count",
    "Maximum",
    "Minimum",
    "Median",
    "Standard Deviation",
}
EXPECTED_OVERLAP = 9
EXPECTED_COLS = 396
EXPECTED_OVERLAP_TICKERS = ["SSFI", "RISR", "FIAX", "HYBI", "QBER", "QBUL", "DUKZ", "FIYY", "KNRG"]


def cell_is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def canonical_cell(value: Any) -> Any:
    if cell_is_blank(value):
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    return value


def canonical_row(row: pd.Series) -> tuple[Any, ...]:
    return tuple(canonical_cell(value) for value in row.tolist())


def read_first_sheet(path: Path) -> tuple[pd.DataFrame, str]:
    xl = pd.ExcelFile(path, engine="calamine")
    sheet_name = xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object, engine="calamine")
    return df, sheet_name


def extract_header_and_product_rows(raw_df: pd.DataFrame) -> tuple[list[Any], pd.DataFrame, int | None]:
    header = raw_df.iloc[0].tolist()
    first_col = raw_df.iloc[:, 0].map(lambda value: "" if cell_is_blank(value) else str(value).strip())
    summary_indices = first_col[first_col.eq(SUMMARY_MARKER)].index.tolist()
    summary_row_idx = summary_indices[0] if summary_indices else None

    end_idx_exclusive = summary_row_idx if summary_row_idx is not None else len(raw_df)
    product = raw_df.iloc[1:end_idx_exclusive].copy()

    # Only drop blank rows directly before Summary Statistics. Do not remove
    # blank rows inside the product area; they are part of the original export
    # positioning and must not change the row count/order unexpectedly.
    while len(product) > 0 and all(cell_is_blank(value) for value in product.iloc[-1].tolist()):
        product = product.iloc[:-1].copy()
    product = product.reset_index(drop=True)
    return header, product, None if summary_row_idx is None else int(summary_row_idx + 1)


def headers_equal(header1: list[Any], header2: list[Any]) -> tuple[bool, list[dict[str, Any]]]:
    max_len = max(len(header1), len(header2))
    diffs: list[dict[str, Any]] = []
    for idx in range(max_len):
        left = canonical_cell(header1[idx]) if idx < len(header1) else ""
        right = canonical_cell(header2[idx]) if idx < len(header2) else ""
        if left != right:
            diffs.append({"列号": idx + 1, "文件1表头": left, "文件2表头": right})
    return not diffs, diffs


def find_longest_continuous_overlap(product1: pd.DataFrame, product2: pd.DataFrame) -> int:
    rows1 = [canonical_row(product1.iloc[i]) for i in range(len(product1))]
    rows2 = [canonical_row(product2.iloc[i]) for i in range(len(product2))]
    max_possible = min(len(rows1), len(rows2))
    for overlap_len in range(max_possible, 0, -1):
        if rows1[-overlap_len:] == rows2[:overlap_len]:
            return overlap_len
    return 0


def make_overlap_examples(product2: pd.DataFrame, overlap_len: int) -> list[dict[str, Any]]:
    examples: list[dict[str, Any]] = []
    for idx in range(min(overlap_len, 20)):
        row = product2.iloc[idx]
        examples.append(
            {
                "序号": idx + 1,
                "Name": row.iloc[0] if len(row) > 0 else "",
                "Ticker": row.iloc[4] if len(row) > 4 else "",
            }
        )
    return examples


def write_output(header: list[Any], merged_product: pd.DataFrame, summary_rows: list[list[Any]]) -> None:
    output_df = merged_product.copy()
    output_df.columns = header
    summary_df = pd.DataFrame(summary_rows, columns=["项目", "内容"])

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, sheet_name=MERGED_SHEET, index=False)
        summary_df.to_excel(writer, sheet_name=SUMMARY_SHEET, index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})
        for sheet_name, df in [(MERGED_SHEET, output_df), (SUMMARY_SHEET, summary_df)]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), max(len(df.columns) - 1, 0))
            for col_idx, col_name in enumerate(df.columns):
                sample_width = len(str(col_name))
                if len(df) > 0:
                    sample = df.iloc[:200, col_idx]
                    sample_width = max(
                        sample_width,
                        int(sample.map(lambda value: 0 if pd.isna(value) else len(str(value))).max()),
                    )
                worksheet.set_column(col_idx, col_idx, min(max(sample_width + 2, 10), 32))
                worksheet.write(0, col_idx, col_name, header_fmt)


def validate_output(
    header: list[Any],
    file1_product_rows: int,
    file2_product_rows: int,
    overlap_len: int,
) -> dict[str, Any]:
    if not OUTPUT_XLSX.exists():
        return {"输出文件是否存在": False}
    xl = pd.ExcelFile(OUTPUT_XLSX, engine="calamine")
    merged = pd.read_excel(OUTPUT_XLSX, sheet_name=MERGED_SHEET, dtype=object, engine="calamine")
    first_col = merged.iloc[:, 0].map(lambda value: "" if cell_is_blank(value) else str(value).strip())
    ticker_col = merged.iloc[:, 4].map(lambda value: "" if cell_is_blank(value) else str(value).strip())

    expected_rows = file1_product_rows + file2_product_rows - overlap_len
    summary_marker_counts = {marker: int(first_col.eq(marker).sum()) for marker in STAT_MARKERS}
    expected_overlap_counts = {ticker: int(ticker_col.eq(ticker).sum()) for ticker in EXPECTED_OVERLAP_TICKERS}

    return {
        "输出文件是否存在": True,
        "输出Sheet是否包含合并Sheet": MERGED_SHEET in xl.sheet_names,
        "输出Sheet是否包含说明Sheet": SUMMARY_SHEET in xl.sheet_names,
        "合并后表头列数": len(merged.columns),
        "合并后表头是否与原始表头完全一致": [canonical_cell(c) for c in merged.columns.tolist()] == [canonical_cell(c) for c in header],
        "是否包含汇总统计行": any(count > 0 for count in summary_marker_counts.values()),
        "汇总统计行命中数量": summary_marker_counts,
        "合并后产品数据行数": len(merged),
        "期望产品数据行数": expected_rows,
        "行数公式是否成立": len(merged) == expected_rows,
        "预期重叠Ticker出现次数": expected_overlap_counts,
    }


def main() -> None:
    if not FILE1.exists():
        raise FileNotFoundError(f"文件1不存在: {FILE1}")
    if not FILE2.exists():
        raise FileNotFoundError(f"文件2不存在: {FILE2}")

    processed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw1, sheet1 = read_first_sheet(FILE1)
    raw2, sheet2 = read_first_sheet(FILE2)
    header1, product1, summary_row1 = extract_header_and_product_rows(raw1)
    header2, product2, summary_row2 = extract_header_and_product_rows(raw2)

    same_header, header_diffs = headers_equal(header1, header2)
    if not same_header:
        diff_lines = [
            "Morningstar 文件合并中止：两个文件表头不一致",
            f"文件1路径: {FILE1}",
            f"文件2路径: {FILE2}",
            "不一致字段:",
            *[
                f"- 列{item['列号']}: 文件1={item['文件1表头']} | 文件2={item['文件2表头']}"
                for item in header_diffs
            ],
        ]
        OUTPUT_LOG.write_text("\n".join(diff_lines) + "\n", encoding="utf-8-sig")
        raise ValueError(f"两个文件表头不一致，已写入日志: {OUTPUT_LOG}")

    overlap_len = find_longest_continuous_overlap(product1, product2)
    overlap_examples = make_overlap_examples(product2, overlap_len)
    merged_product = pd.concat([product1, product2.iloc[overlap_len:]], ignore_index=True)

    summary_rows = [
        ["文件1路径", str(FILE1)],
        ["文件2路径", str(FILE2)],
        ["文件1Sheet", sheet1],
        ["文件2Sheet", sheet2],
        ["文件1原始行数", raw1.shape[0]],
        ["文件2原始行数", raw2.shape[0]],
        ["文件1原始列数", raw1.shape[1]],
        ["文件2原始列数", raw2.shape[1]],
        ["文件1产品数据行数", len(product1)],
        ["文件2产品数据行数", len(product2)],
        ["识别到的连续重叠行数", overlap_len],
        ["合并后产品数据行数", len(merged_product)],
        ["是否删除 Summary Statistics 区", "是"],
        ["是否全局去重", "否"],
        ["处理时间", processed_at],
    ]
    write_output(header1, merged_product, summary_rows)
    validation = validate_output(header1, len(product1), len(product2), overlap_len)

    overlap_warning = ""
    if overlap_len != EXPECTED_OVERLAP:
        overlap_warning = f"醒目标注：实际识别重叠行数不是9，而是 {overlap_len}，需要人工复核。"

    log_lines = [
        "Morningstar 美国ETF 2010-2026 合并日志",
        "=" * 50,
        f"文件1路径: {FILE1}",
        f"文件2路径: {FILE2}",
        f"文件1原始行数、列数: {raw1.shape[0]}, {raw1.shape[1]}",
        f"文件2原始行数、列数: {raw2.shape[0]}, {raw2.shape[1]}",
        f"表头是否完全一致: {'是' if same_header else '否'}",
        f"文件1产品数据区行数: {len(product1)}",
        f"文件2产品数据区行数: {len(product2)}",
        f"自动识别出的连续重叠行数: {overlap_len}",
        overlap_warning,
        "重叠行前20条示例:",
        *[f"- {item['序号']}: Name={item['Name']}; Ticker={item['Ticker']}" for item in overlap_examples],
        f"合并后产品数据行数: {len(merged_product)}",
        "是否删除 Summary Statistics 汇总区: 是",
        "是否进行全局去重: 否",
        f"输出文件路径: {OUTPUT_XLSX}",
        f"处理完成时间: {processed_at}",
        "",
        "自动校验结果:",
        *[f"- {key}: {value}" for key, value in validation.items()],
    ]
    OUTPUT_LOG.write_text("\n".join(line for line in log_lines if line != "") + "\n", encoding="utf-8-sig")

    print("合并成功")
    print(f"文件1产品数据行数: {len(product1)}")
    print(f"文件2产品数据行数: {len(product2)}")
    print(f"连续重叠行数: {overlap_len}")
    print(f"合并后总行数: {len(merged_product)}")
    print(f"输出文件路径: {OUTPUT_XLSX}")
    if overlap_warning:
        print(overlap_warning)


if __name__ == "__main__":
    main()
