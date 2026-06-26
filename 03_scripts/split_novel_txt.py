from __future__ import annotations

import hashlib
import re
import sys
from dataclasses import dataclass
from pathlib import Path


SOURCE = Path(r"C:\Users\lvdon\xwechat_files\wxid_0tl2ozelncw822_0d86\msg\file\2026-06\重生之都市修仙.txt")
OUT_DIR = SOURCE.with_name("重生之都市修仙_拆分版")
PART1 = OUT_DIR / "重生之都市修仙_第1部分.txt"
PART2 = OUT_DIR / "重生之都市修仙_第2部分.txt"
REPORT = OUT_DIR / "拆分说明.txt"

MAX_BYTES = 50 * 1024 * 1024

CN_NUM = "零〇一二三四五六七八九十百千万亿两壹贰叁肆伍陆柒捌玖拾佰仟"


@dataclass(frozen=True)
class DecodedText:
    text: str
    encoding_label: str
    encode_name: str
    bom_len: int


@dataclass(frozen=True)
class Boundary:
    char_index: int
    byte_offset: int
    title: str
    line_no: int
    kind: str


def fail(message: str) -> None:
    print(f"失败：{message}")
    sys.exit(1)


def format_size(num_bytes: int) -> str:
    return f"{num_bytes / (1024 * 1024):.2f} MB ({num_bytes:,} bytes)"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def decode_score(text: str) -> tuple[int, int, int, int]:
    sample = text[:200_000]
    cjk = sum(1 for ch in sample if "\u4e00" <= ch <= "\u9fff")
    punct = sum(1 for ch in sample if ch in "，。！？；：“”‘’、（）《》")
    controls = sum(1 for ch in sample if ord(ch) < 32 and ch not in "\r\n\t")
    mojibake = sample.count("锟") + sample.count("�") + sample.count("Ã") + sample.count("Â")
    return (cjk + punct * 2 - controls * 100 - mojibake * 500, cjk, punct, -controls)


def decode_bytes(raw: bytes) -> DecodedText:
    candidates: list[tuple[str, str, int]] = []
    if raw.startswith(b"\xef\xbb\xbf"):
        candidates.append(("UTF-8 with BOM", "utf-8", 3))
    candidates.extend(
        [
            ("UTF-8", "utf-8", 0),
            ("GB18030", "gb18030", 0),
            ("GBK", "gbk", 0),
            ("ANSI/CP936", "cp936", 0),
        ]
    )

    decoded: list[tuple[tuple[int, int, int, int], DecodedText]] = []
    seen: set[tuple[str, int]] = set()
    for label, enc, bom_len in candidates:
        key = (enc, bom_len)
        if key in seen:
            continue
        seen.add(key)
        try:
            body = raw[bom_len:]
            text = body.decode(enc, errors="strict")
        except UnicodeDecodeError:
            continue
        decoded.append((decode_score(text), DecodedText(text, label, enc, bom_len)))

    if not decoded:
        fail("无法用 UTF-8、GB18030、GBK 或 ANSI/CP936 严格解码原文件。")

    decoded.sort(key=lambda item: item[0], reverse=True)
    return decoded[0][1]


CHAPTER_RE = re.compile(
    rf"""
    ^[\s　]*(?:正文[\s　]*)?
    (?:
        第[\s　]*[0-9０-９{CN_NUM}]{{1,12}}[\s　]*(?:章|章节|节|回|卷|部|集|篇)
        |
        (?:卷|篇|部)[\s　]*[0-9０-９{CN_NUM}]{{1,12}}
        |
        第[\s　]*[0-9０-９{CN_NUM}]{{1,12}}[\s　]*(?:卷|部|篇)
    )
    (?:[\s　:：、.\-－—].*)?
    $""",
    re.VERBOSE,
)

STAGE_KEYWORDS = (
    "都市",
    "地球",
    "北琼",
    "昆墟",
    "天荒",
    "宇宙",
    "星海",
    "星河",
    "中央星河",
    "仙界",
    "人间",
    "飞升",
)


def is_chapter_title(line: str) -> bool:
    compact = line.strip()
    if not compact or len(compact) > 90:
        return False
    return CHAPTER_RE.match(compact) is not None


def is_stage_title(line: str) -> bool:
    compact = line.strip()
    if not compact or len(compact) > 40:
        return False
    if any(mark in compact for mark in "，,。！？!?；;：“”‘’\""):
        return False
    stage_core = "|".join(map(re.escape, STAGE_KEYWORDS))
    stage_re = re.compile(
        rf"^(?:第[\s　]*[0-9０-９{CN_NUM}]{{1,12}}[\s　]*(?:篇|卷|部)[\s　:：、.\-－—]*)?"
        rf"(?:{stage_core})(?:篇|卷|部)?$"
    )
    if stage_re.match(compact):
        return True
    titled_stage_re = re.compile(
        rf"^(?:篇|卷|部)[\s　]*[0-9０-９{CN_NUM}]{{1,12}}[\s　:：、.\-－—]*(?:{stage_core})(?:篇|卷|部)?$"
    )
    if titled_stage_re.match(compact):
        return True
    if not re.match(
        rf"^第[\s　]*[0-9０-９{CN_NUM}]{{1,12}}[\s　]*(?:篇|卷|部)[\s　:：、.\-－—]*",
        compact,
    ):
        return False
    return any(keyword in compact for keyword in STAGE_KEYWORDS)


def line_records(text: str) -> list[tuple[int, int, str]]:
    records: list[tuple[int, int, str]] = []
    line_no = 1
    pos = 0
    for line in text.splitlines(keepends=True):
        stripped = line.rstrip("\r\n")
        records.append((pos, line_no, stripped))
        pos += len(line)
        line_no += 1
    if pos < len(text):
        records.append((pos, line_no, text[pos:]))
    return records


def offsets_for_indices(text: str, indices: list[int], encode_name: str, bom_len: int) -> dict[int, int]:
    offsets: dict[int, int] = {}
    prev_index = 0
    current_offset = bom_len
    for index in sorted(set(indices)):
        if index < prev_index:
            continue
        current_offset += len(text[prev_index:index].encode(encode_name))
        offsets[index] = current_offset
        prev_index = index
    return offsets


def find_boundaries(decoded: DecodedText) -> tuple[list[Boundary], list[Boundary]]:
    chapters_raw: list[tuple[int, str, int]] = []
    stages_raw: list[tuple[int, str, int]] = []

    for char_index, line_no, line in line_records(decoded.text):
        if is_chapter_title(line):
            chapters_raw.append((char_index, line.strip(), line_no))
        elif is_stage_title(line):
            stages_raw.append((char_index, line.strip(), line_no))

    all_indices = [item[0] for item in chapters_raw] + [item[0] for item in stages_raw]
    offsets = offsets_for_indices(decoded.text, all_indices, decoded.encode_name, decoded.bom_len)

    chapters = [
        Boundary(char_index, offsets[char_index], title, line_no, "chapter")
        for char_index, title, line_no in chapters_raw
    ]
    stages = [
        Boundary(char_index, offsets[char_index], title, line_no, "stage")
        for char_index, title, line_no in stages_raw
    ]
    return chapters, stages


def previous_chapter(chapters: list[Boundary], split_char_index: int) -> Boundary | None:
    prev: Boundary | None = None
    for chapter in chapters:
        if chapter.char_index < split_char_index:
            prev = chapter
        else:
            break
    return prev


def next_chapter(chapters: list[Boundary], split_char_index: int) -> Boundary | None:
    for chapter in chapters:
        if chapter.char_index >= split_char_index:
            return chapter
    return None


def choose_split(raw_size: int, chapters: list[Boundary], stages: list[Boundary]) -> tuple[Boundary, bool, str]:
    if raw_size >= 2 * MAX_BYTES:
        fail(
            "原文件大小达到或超过两个 50MB 文件的总容量上限，无法拆成两个都小于 50MB 的文件；建议拆成 3 个或更多文件。"
        )

    if len(chapters) < 2:
        fail("未识别到足够章节标题，无法保证不在章节中间硬切。")

    def valid(byte_offset: int) -> bool:
        return 0 < byte_offset < raw_size and byte_offset < MAX_BYTES and (raw_size - byte_offset) < MAX_BYTES

    target = raw_size / 2

    valid_stages = [stage for stage in stages if valid(stage.byte_offset)]
    if valid_stages:
        chosen = min(valid_stages, key=lambda item: abs(item.byte_offset - target))
        return chosen, True, "检测到大篇章结构，采用最接近中间且满足大小限制的大篇章节点拆分。"

    valid_chapters = [chapter for chapter in chapters[1:] if valid(chapter.byte_offset)]
    if not valid_chapters:
        fail("虽然原文件小于 100MB，但没有找到能让两个输出文件都小于 50MB 的章节边界。")

    chosen = min(valid_chapters, key=lambda item: abs(item.byte_offset - target))
    return chosen, False, "未检测到可用的大篇章拆分点，按章节边界和文件大小自动平衡拆分。"


def chapter_range(chapters: list[Boundary], start_char: int, end_char: int | None) -> str:
    selected = [
        chapter
        for chapter in chapters
        if chapter.char_index >= start_char and (end_char is None or chapter.char_index < end_char)
    ]
    if not selected:
        return "未识别到章节标题"
    if len(selected) == 1:
        return selected[0].title
    return f"{selected[0].title} 至 {selected[-1].title}"


def validate_outputs(raw: bytes, split: Boundary, chapters: list[Boundary]) -> None:
    if not PART1.exists() or not PART2.exists():
        fail("输出文件不存在。")
    part1 = PART1.read_bytes()
    part2 = PART2.read_bytes()
    if len(part1) >= MAX_BYTES or len(part2) >= MAX_BYTES:
        fail(
            f"输出文件大小超限：第1部分 {format_size(len(part1))}，第2部分 {format_size(len(part2))}。"
        )
    if part1 + part2 != raw:
        fail("两个输出文件拼接后的字节内容与原文件不完全一致。")

    chapter_offsets = {chapter.byte_offset for chapter in chapters}
    if split.kind == "chapter" and split.byte_offset not in chapter_offsets:
        fail("拆分点不是已识别章节边界。")
    if split.kind == "stage":
        prev = previous_chapter(chapters, split.char_index)
        nxt = next_chapter(chapters, split.char_index)
        if prev is None or nxt is None:
            fail("大篇章拆分点不在两个已识别章节之间，无法确认章节完整性。")


def main() -> None:
    if not SOURCE.exists():
        fail(f"原文件不存在：{SOURCE}")
    if not SOURCE.is_file():
        fail(f"原路径不是文件：{SOURCE}")

    raw = SOURCE.read_bytes()
    raw_size = len(raw)
    decoded = decode_bytes(raw)
    chapters, stages = find_boundaries(decoded)
    split, used_stage, split_reason = choose_split(raw_size, chapters, stages)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    PART1.write_bytes(raw[: split.byte_offset])
    PART2.write_bytes(raw[split.byte_offset :])

    validate_outputs(raw, split, chapters)

    prev = previous_chapter(chapters, split.char_index)
    next_ch = next_chapter(chapters, split.char_index)
    part1_size = PART1.stat().st_size
    part2_size = PART2.stat().st_size

    detected_stage_text = "是" if stages else "否"
    used_stage_text = "是" if used_stage else "否"
    split_after = prev.title if prev else "原文开头之前"
    split_before = next_ch.title if next_ch else "原文结尾"

    report = "\n".join(
        [
            "拆分说明",
            "",
            f"原始文件：{SOURCE}",
            f"原始文件大小：{format_size(raw_size)}",
            f"使用的文本编码：{decoded.encoding_label}",
            f"原文件 SHA256：{sha256_bytes(raw)}",
            f"总章节数：{len(chapters)}",
            f"检测到的大篇章候选数：{len(stages)}",
            f"是否检测到了“都市篇 / 宇宙篇”等大篇章结构：{detected_stage_text}",
            f"是否采用大篇章节点拆分：{used_stage_text}",
            f"拆分策略说明：{split_reason}",
            "",
            f"第1部分文件：{PART1}",
            f"第1部分大小：{format_size(part1_size)}",
            f"第1部分包含的章节范围：{chapter_range(chapters, 0, split.char_index)}",
            "",
            f"第2部分文件：{PART2}",
            f"第2部分大小：{format_size(part2_size)}",
            f"第2部分包含的章节范围：{chapter_range(chapters, split.char_index, None)}",
            "",
            f"实际拆分点：在“{split_after}”之后拆分",
            f"拆分点后一章/节点：{split.title if split.kind == 'stage' else split_before}",
            f"拆分点类型：{'大篇章标题' if split.kind == 'stage' else '章节标题'}",
            f"拆分点行号：{split.line_no}",
            f"拆分点字节位置：{split.byte_offset:,}",
            "",
            "校验结果：",
            "两个输出文件都存在：是",
            "两个输出文件都小于 50MB：是",
            "两个文件拼接后与原文字节完全一致：是",
            "没有章节被截断：是",
            "",
        ]
    )
    REPORT.write_text(report, encoding="utf-8-sig", newline="\r\n")

    print("验收结果：成功")
    print(f"原始文件大小：{format_size(raw_size)}")
    print(f"使用编码：{decoded.encoding_label}")
    print(f"总章节数：{len(chapters)}")
    print(f"大篇章候选数：{len(stages)}，采用大篇章拆分：{used_stage_text}")
    print(f"拆分点：在“{split_after}”之后")
    print(f"第1部分：{format_size(part1_size)}")
    print(f"第2部分：{format_size(part2_size)}")
    print(f"输出目录：{OUT_DIR}")
    print(f"说明文件：{REPORT}")


if __name__ == "__main__":
    main()
