#!/usr/bin/env python3
"""
PDF 轉 EPUB / KEPUB 轉換工具

使用 PyMuPDF 解析 PDF 文字區塊，清理頁首/頁尾/頁碼，
依字體大小辨識標題層級，再以 ebooklib 封裝成 EPUB，
最後透過 Calibre ebook-convert 轉成 Kobo 友善的 KEPUB。
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import subprocess
import zipfile
from xml.etree import ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import fitz  # PyMuPDF
from ebooklib import epub


# ---------------------------------------------------------------------------
# 資料結構
# ---------------------------------------------------------------------------


@dataclass
class TextBlock:
    """代表 PDF 頁面上的一個文字區塊。"""

    page_index: int
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    font_size: float = 0.0
    is_heading: bool = False
    heading_level: int = 0  # 0 = 段落, 1~3 = h1~h3


@dataclass
class ConversionConfig:
    """轉換過程的可調參數。"""

    header_ratio: float = 0.08  # 頁面上方視為頁首的比例
    footer_ratio: float = 0.08  # 頁面下方視為頁尾的比例
    min_block_chars: int = 2  # 過短區塊直接忽略
    title: str = ""
    author: str = "Unknown"
    language: str = "zh"


# 常見頁碼格式（獨立一行或極短文字）
PAGE_NUMBER_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"^\d{1,4}$"),
    re.compile(r"^第\s*\d+\s*頁$"),
    re.compile(r"^page\s*\d+\s*$", re.IGNORECASE),
    re.compile(r"^-\s*\d+\s*-$"),
    re.compile(r"^\[\s*\d+\s*\]$"),
    re.compile(r"^\d+\s*/\s*\d+$"),
    re.compile(r"^—\s*\d+\s*—$"),
)


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------


def _collect_span_font_sizes(page: fitz.Page) -> list[tuple[fitz.Rect, float]]:
    """從頁面 dict 結構收集每個 span 的邊界框與字體大小。"""
    spans: list[tuple[fitz.Rect, float]] = []
    page_dict = page.get_text("dict")

    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue
                bbox = fitz.Rect(span["bbox"])
                size = float(span.get("size", 0))
                spans.append((bbox, size))

    return spans


def _estimate_block_font_size(block_rect: fitz.Rect, spans: list[tuple[fitz.Rect, float]]) -> float:
    """依區塊與 span 的重疊面積，估算該區塊的代表字體大小。"""
    weighted_sizes: list[tuple[float, float]] = []

    for span_rect, size in spans:
        intersection = block_rect & span_rect
        if intersection.is_empty:
            continue
        overlap_area = intersection.get_area()
        if overlap_area > 0 and size > 0:
            weighted_sizes.append((size, overlap_area))

    if not weighted_sizes:
        return 0.0

    total_area = sum(area for _, area in weighted_sizes)
    if total_area <= 0:
        return 0.0

    return sum(size * area for size, area in weighted_sizes) / total_area


def extract_cover_image(pdf_path: Path, cover_path: Path) -> None:
    """將 PDF 第一頁（第 0 頁）渲染為封面圖片。"""
    with fitz.open(pdf_path) as doc:
        if len(doc) == 0:
            return

        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        cover_path.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(cover_path))


def extract_blocks_from_pdf(pdf_path: Path) -> list[TextBlock]:
    """
    使用 page.get_text('blocks') 讀取 PDF，並補充字體大小資訊。

    blocks 格式：(x0, y0, x1, y1, text, block_no, block_type)
    block_type == 0 表示文字區塊。
    第一頁作為封面，內文從第二頁開始解析。
    """
    blocks: list[TextBlock] = []

    with fitz.open(pdf_path) as doc:
        for page_index in range(1, len(doc)):
            page = doc[page_index]
            span_fonts = _collect_span_font_sizes(page)
            raw_blocks = page.get_text("blocks")

            for item in raw_blocks:
                if len(item) < 7:
                    continue

                x0, y0, x1, y1, text, _block_no, block_type = item[:7]
                if block_type != 0:
                    continue

                cleaned = _normalize_whitespace(str(text))
                if not cleaned:
                    continue

                block_rect = fitz.Rect(x0, y0, x1, y1)
                font_size = _estimate_block_font_size(block_rect, span_fonts)

                blocks.append(
                    TextBlock(
                        page_index=page_index,
                        x0=x0,
                        y0=y0,
                        x1=x1,
                        y1=y1,
                        text=cleaned,
                        font_size=font_size,
                    )
                )

    return blocks


def _normalize_whitespace(text: str) -> str:
    """合併多餘空白，保留段落內換行為單一空格。"""
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    return " ".join(lines)


# ---------------------------------------------------------------------------
# 結構清理
# ---------------------------------------------------------------------------


def _looks_like_page_number(text: str) -> bool:
    """判斷文字是否像頁碼。"""
    candidate = text.strip()
    if not candidate:
        return False
    return any(pattern.match(candidate) for pattern in PAGE_NUMBER_PATTERNS)


def _is_header_or_footer(block: TextBlock, page_height: float, config: ConversionConfig) -> bool:
    """依 Y 座標與頁碼特徵過濾頁首、頁尾。"""
    header_limit = page_height * config.header_ratio
    footer_limit = page_height * (1.0 - config.footer_ratio)

    in_header = block.y1 <= header_limit
    in_footer = block.y0 >= footer_limit

    if not (in_header or in_footer):
        return False

    # 頁首/頁尾區域內：頁碼或極短文字通常可移除
    if _looks_like_page_number(block.text):
        return True

    if len(block.text) <= 40 and (in_header or in_footer):
        return True

    return False


def filter_header_footer(blocks: list[TextBlock], page_heights: dict[int, float], config: ConversionConfig) -> list[TextBlock]:
    """移除頁首、頁尾與疑似頁碼區塊。"""
    kept: list[TextBlock] = []

    for block in blocks:
        page_height = page_heights.get(block.page_index, 0.0)
        if page_height <= 0:
            kept.append(block)
            continue

        if len(block.text) < config.min_block_chars:
            continue

        if _is_header_or_footer(block, page_height, config):
            continue

        kept.append(block)

    return kept


def _detect_body_font_size(font_sizes: Iterable[float]) -> float:
    """以眾數（最常見字體大小）作為內文字級基準。"""
    rounded = [round(size, 1) for size in font_sizes if size > 0]
    if not rounded:
        return 12.0

    counter = Counter(rounded)
    return counter.most_common(1)[0][0]


def classify_headings(blocks: list[TextBlock]) -> None:
    """
    依字體大小將區塊標記為標題或段落。

    大於內文基準一定比例者，依大小分為 h1~h3。
    """
    sizes = [block.font_size for block in blocks if block.font_size > 0]
    if not sizes:
        return

    body_size = _detect_body_font_size(sizes)
    heading_candidates: list[tuple[TextBlock, float]] = []

    for block in blocks:
        if block.font_size <= body_size * 1.08:
            continue
        # 過長文字較不像標題
        if len(block.text) > 120:
            continue
        heading_candidates.append((block, block.font_size))

    if not heading_candidates:
        return

    unique_sizes = sorted({round(size, 1) for _, size in heading_candidates}, reverse=True)
    size_to_level = {size: min(index + 1, 3) for index, size in enumerate(unique_sizes)}

    for block, size in heading_candidates:
        rounded = round(size, 1)
        block.is_heading = True
        block.heading_level = size_to_level.get(rounded, 3)


def merge_adjacent_paragraphs(blocks: list[TextBlock]) -> list[TextBlock]:
    """
    合併相鄰、同層級的段落區塊（同一頁、垂直距離近、皆為正文）。

    有助於還原 get_text('blocks') 可能切太細的段落。
    """
    if not blocks:
        return []

    merged: list[TextBlock] = [blocks[0]]

    for current in blocks[1:]:
        previous = merged[-1]
        same_page = current.page_index == previous.page_index
        both_paragraph = not previous.is_heading and not current.is_heading
        close_vertically = current.y0 - previous.y1 < 18

        if same_page and both_paragraph and close_vertically:
            previous.text = f"{previous.text} {current.text}"
            previous.y1 = max(previous.y1, current.y1)
            previous.x1 = max(previous.x1, current.x1)
        else:
            merged.append(current)

    return merged


def clean_and_structure_blocks(blocks: list[TextBlock], pdf_path: Path, config: ConversionConfig) -> list[TextBlock]:
    """執行完整的結構清理流程。"""
    page_heights: dict[int, float] = {}

    with fitz.open(pdf_path) as doc:
        for index in range(1, len(doc)):
            page_heights[index] = doc[index].rect.height

    filtered = filter_header_footer(blocks, page_heights, config)
    classify_headings(filtered)
    return merge_adjacent_paragraphs(filtered)


# ---------------------------------------------------------------------------
# EPUB 封裝
# ---------------------------------------------------------------------------


def _escape_html(text: str) -> str:
    return html.escape(text, quote=False)


def blocks_to_html(blocks: list[TextBlock]) -> str:
    """將文字區塊序列化為 HTML 片段。"""
    parts: list[str] = []

    for block in blocks:
        content = _escape_html(block.text)
        if block.is_heading and block.heading_level > 0:
            level = block.heading_level
            parts.append(f"<h{level}>{content}</h{level}>")
        else:
            parts.append(f"<p>{content}</p>")

    body = "\n".join(parts)
    return f"""<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh">
<head>
  <meta charset="utf-8"/>
  <title>Document</title>
  <style>
    body {{ font-family: serif; line-height: 1.6; margin: 1em; }}
    h1, h2, h3 {{ margin-top: 1.2em; margin-bottom: 0.6em; }}
    p {{ text-indent: 2em; margin: 0.6em 0; }}
  </style>
</head>
<body>
{body}
</body>
</html>
"""


def _split_blocks_into_chapters(blocks: list[TextBlock], blocks_per_chapter: int = 120) -> list[list[TextBlock]]:
    """
    依 h1 標題或固定區塊數切分章節。

    若存在 h1，優先以其分章；否則按區塊數分批。
    """
    h1_indices = [i for i, block in enumerate(blocks) if block.is_heading and block.heading_level == 1]

    if h1_indices:
        chapters: list[list[TextBlock]] = []
        for idx, start in enumerate(h1_indices):
            end = h1_indices[idx + 1] if idx + 1 < len(h1_indices) else len(blocks)
            chapters.append(blocks[start:end])
        if h1_indices[0] > 0:
            chapters.insert(0, blocks[: h1_indices[0]])
        return [chapter for chapter in chapters if chapter]

    return [blocks[i : i + blocks_per_chapter] for i in range(0, len(blocks), blocks_per_chapter)]


def build_epub(
    blocks: list[TextBlock],
    output_path: Path,
    config: ConversionConfig,
    cover_path: Path | None = None,
) -> None:
    """使用 ebooklib 建立 EPUB 檔案。"""
    book = epub.EpubBook()
    book.set_identifier(f"pdf2epub-{output_path.stem}")
    book.set_title(config.title or output_path.stem)
    book.set_language(config.language)
    book.add_author(config.author)

    cover_page_id: str | None = None
    if cover_path and cover_path.is_file():
        with open(cover_path, "rb") as cover_file:
            book.set_cover("cover.png", cover_file.read())

        cover_page = book.get_item_with_id("cover")
        if cover_page is not None:
            cover_page.is_linear = True
            cover_page_id = "cover"

    chapters = _split_blocks_into_chapters(blocks)
    spine: list[epub.EpubItem | str] = []
    if cover_page_id:
        spine.append(cover_page_id)
    spine.append("nav")
    toc: list[epub.Link] = []

    for index, chapter_blocks in enumerate(chapters, start=1):
        chapter_id = f"chapter_{index:03d}"
        chapter_file = f"{chapter_id}.xhtml"

        # 章節標題：優先取第一個標題，否則用預設名稱
        first_heading = next((b.text for b in chapter_blocks if b.is_heading), None)
        chapter_title = first_heading or f"第 {index} 章"

        chapter = epub.EpubHtml(
            title=chapter_title[:120],
            file_name=chapter_file,
            lang=config.language,
        )
        chapter.content = blocks_to_html(chapter_blocks)
        book.add_item(chapter)

        spine.append(chapter)
        toc.append(epub.Link(chapter_file, chapter_title[:120], chapter_id))

    book.toc = toc
    book.spine = spine
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())

    output_path.parent.mkdir(parents=True, exist_ok=True)
    epub.write_epub(str(output_path), book, {})


# ---------------------------------------------------------------------------
# Calibre KEPUB 轉換
# ---------------------------------------------------------------------------


def _resolve_ebook_convert() -> str:
    """尋找 Calibre 的 ebook-convert 可執行檔。"""
    found = shutil.which("ebook-convert")
    if found:
        return found

    # macOS 常見安裝路徑（App bundle 未加入 PATH 時）
    mac_candidates = (
        "/Applications/calibre.app/Contents/MacOS/ebook-convert",
        "/usr/local/bin/ebook-convert",
        "/opt/homebrew/bin/ebook-convert",
    )
    for candidate in mac_candidates:
        if Path(candidate).is_file():
            return candidate

    raise FileNotFoundError(
        "找不到 ebook-convert。請安裝 Calibre，並確保 CLI 可在終端機執行 "
        "（macOS 可將 /Applications/calibre.app/Contents/MacOS 加入 PATH）。"
    )


def convert_epub_to_kepub(epub_path: Path) -> Path:
    """
    使用 Calibre ebook-convert 將 EPUB 轉成 Kobo KEPUB。

    經 Calibre 轉換管線可重整 HTML/CSS 結構，產出 XXXX.kepub.epub，
    降低 Kobo 韌體解析不良標記而卡死的風險。
    """
    if not epub_path.is_file():
        raise FileNotFoundError(f"找不到 EPUB 檔案：{epub_path}")

    ebook_convert = _resolve_ebook_convert()
    kepub_path = epub_path.with_name(f"{epub_path.stem}.kepub.epub")

    cmd = [
        ebook_convert,
        str(epub_path),
        str(kepub_path),
        "--output-profile=kobo",
        "--epub-version=3",
        "--flow-size=0",
        "--no-svg-cover",
        "--pretty-print",
    ]

    print(f"      執行：{' '.join(cmd)}")
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(
            f"ebook-convert 轉換失敗（結束碼 {result.returncode}）"
            + (f"：\n{detail}" if detail else "")
        )

    if not kepub_path.is_file():
        raise RuntimeError(f"ebook-convert 已結束，但找不到輸出檔：{kepub_path}")

    return kepub_path


# ---------------------------------------------------------------------------
# 作者資料夾歸檔
# ---------------------------------------------------------------------------


INVALID_FILENAME_CHARS = r'[<>:"/\\|?*\x00-\x1F]'


def _normalize_author_name(author: str) -> str:
    """
    作者名稱正規化：
    - 清理空白
    - 統一首字母大寫（忽略原始大小寫）
    """
    cleaned = _normalize_whitespace(author)
    if not cleaned:
        return "Unknown"
    return cleaned.title()


def _safe_folder_name(name: str) -> str:
    """移除不適合作為資料夾名稱的字元。"""
    safe = re.sub(INVALID_FILENAME_CHARS, "_", name).strip(" .")
    return safe or "Unknown"


def _extract_author_from_filename(path: Path) -> str | None:
    """
    從檔名推測作者，例如：
    - Author - Title.epub
    - Author_ Title.epub
    """
    stem = path.stem
    for sep in (" - ", "_-_", "_", "-"):
        if sep in stem:
            candidate = stem.split(sep, 1)[0].strip()
            if len(candidate) >= 2:
                return candidate
    return None


def _extract_author_from_epub_metadata(epub_path: Path) -> str | None:
    """
    從 EPUB/KEPUB 的 OPF metadata 讀取作者（dc:creator）。
    透過 zip + XML 解析，避免依賴特定 reader 實作。
    """
    try:
        with zipfile.ZipFile(epub_path, "r") as archive:
            container_xml = archive.read("META-INF/container.xml")
            container_root = ET.fromstring(container_xml)
            rootfile = container_root.find(".//{*}rootfile")
            if rootfile is None:
                return None

            opf_path = rootfile.attrib.get("full-path", "").strip()
            if not opf_path:
                return None

            opf_xml = archive.read(opf_path)
            opf_root = ET.fromstring(opf_xml)
            creators = opf_root.findall(".//{*}metadata/{*}creator")
            for creator in creators:
                if creator.text and creator.text.strip():
                    return creator.text.strip()
    except (KeyError, ET.ParseError, OSError, zipfile.BadZipFile):
        return None

    return None


def detect_author_name(epub_path: Path) -> str:
    """
    依序由 metadata、檔名抓作者；若都失敗，回傳 Unknown。
    """
    author_from_meta = _extract_author_from_epub_metadata(epub_path)
    if author_from_meta:
        return _normalize_author_name(author_from_meta)

    author_from_name = _extract_author_from_filename(epub_path)
    if author_from_name:
        return _normalize_author_name(author_from_name)

    return "Unknown"


def copy_kepub_to_author_folder(kepub_path: Path, base_dir: Path | None = None) -> Path:
    """
    將 .kepub.epub 複製到作者資料夾。
    預設建立在 kepub 檔案同層目錄下。
    """
    if not kepub_path.is_file():
        raise FileNotFoundError(f"找不到 KEPUB 檔案：{kepub_path}")

    target_base = base_dir or kepub_path.parent
    author_name = detect_author_name(kepub_path)
    author_folder = target_base / _safe_folder_name(author_name)
    author_folder.mkdir(parents=True, exist_ok=True)

    target_file = author_folder / kepub_path.name
    shutil.copy2(kepub_path, target_file)
    return target_file


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def convert_pdf_to_epub(
    pdf_path: Path,
    epub_path: Path,
    config: ConversionConfig | None = None,
    *,
    skip_kepub: bool = False,
) -> None:
    """將 PDF 轉換為 EPUB，並預設再轉成 KEPUB。"""
    config = config or ConversionConfig()

    if not pdf_path.is_file():
        raise FileNotFoundError(f"找不到 PDF 檔案：{pdf_path}")

    if not config.title:
        config.title = pdf_path.stem

    cover_path = epub_path.parent / "cover.png"
    total_steps = 4 if skip_kepub else 6

    try:
        print(f"[1/{total_steps}] 提取封面：{pdf_path} 第 1 頁")
        extract_cover_image(pdf_path, cover_path)
        if cover_path.is_file():
            print(f"      已儲存封面：{cover_path}")
        else:
            print("      PDF 無頁面，略過封面")

        print(f"[2/{total_steps}] 解析 PDF 內文：{pdf_path}")
        raw_blocks = extract_blocks_from_pdf(pdf_path)
        print(f"      讀取 {len(raw_blocks)} 個文字區塊（已跳過封面頁）")

        print(f"[3/{total_steps}] 清理結構並辨識標題")
        structured_blocks = clean_and_structure_blocks(raw_blocks, pdf_path, config)
        heading_count = sum(1 for block in structured_blocks if block.is_heading)
        print(f"      保留 {len(structured_blocks)} 個區塊，其中標題 {heading_count} 個")

        if not structured_blocks:
            raise ValueError("未從 PDF 提取到可用文字，無法建立 EPUB")

        print(f"[4/{total_steps}] 封裝 EPUB：{epub_path}")
        build_epub(structured_blocks, epub_path, config, cover_path=cover_path)
        print(f"      已產生：{epub_path}")

        if skip_kepub:
            print("轉換完成（已略過 KEPUB）。")
            return

        print(f"[5/{total_steps}] 以 Calibre 轉成 KEPUB")
        kepub_path = convert_epub_to_kepub(epub_path)
        print(f"      已產生：{kepub_path}")

        print(f"[6/{total_steps}] 依作者自動歸檔 KEPUB")
        copied_path = copy_kepub_to_author_folder(kepub_path)
        print(f"      已複製至：{copied_path}")
        print("轉換完成。")
    finally:
        if cover_path.is_file():
            cover_path.unlink()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="將 PDF 轉為 EPUB，並以 Calibre 再轉成 Kobo KEPUB",
    )
    parser.add_argument("input_pdf", type=Path, help="輸入 PDF 檔案路徑")
    parser.add_argument("output_epub", type=Path, help="輸出 EPUB 檔案路徑")
    parser.add_argument("--title", default="", help="電子書標題（預設使用 PDF 檔名）")
    parser.add_argument("--author", default="Unknown", help="作者名稱")
    parser.add_argument("--language", default="zh", help="語言代碼，例如 zh、en")
    parser.add_argument(
        "--header-ratio",
        type=float,
        default=0.08,
        help="頁首區域高度比例（0~1，預設 0.08）",
    )
    parser.add_argument(
        "--footer-ratio",
        type=float,
        default=0.08,
        help="頁尾區域高度比例（0~1，預設 0.08）",
    )
    parser.add_argument(
        "--skip-kepub",
        action="store_true",
        help="只輸出 EPUB，不呼叫 ebook-convert 轉成 KEPUB",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    config = ConversionConfig(
        title=args.title,
        author=args.author,
        language=args.language,
        header_ratio=args.header_ratio,
        footer_ratio=args.footer_ratio,
    )

    convert_pdf_to_epub(
        args.input_pdf,
        args.output_epub,
        config,
        skip_kepub=args.skip_kepub,
    )


if __name__ == "__main__":
    main()
