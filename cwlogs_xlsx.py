#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CloudWatch Logs の内容を Meiryo UI の xlsx に書く。

RHEL 9.8 の Python 3.9 標準ライブラリだけで xlsx (Office Open XML) を作る。
openpyxl などの追加インストールは使わない。フォントファイルは埋め込まない。
ブックを開く Windows 側に Meiryo UI がある前提で、書式にフォント名だけを記録する。
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Iterator, Optional, Sequence
from xml.sax.saxutils import escape as _xml_escape

# Excel のシート行上限。バナー行を残すため、データに使える最終行はこれ未満にする。
EXCEL_MAX_ROW = 1048576
# セルの文字数上限。超えるログ行は複数行に分割する。
EXCEL_MAX_CELL_CHARS = 32767

FONT_NAME = "Meiryo UI"

# 白背景の上で判別できる強調色。指定文言の出現順に割り当て、足りなければ繰り返す。
HIGHLIGHT_COLORS = (
    "C0392B",
    "D35400",
    "1A5276",
    "6C3483",
    "0E6655",
    "A15C07",
)

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_NS_DC = "http://purl.org/dc/elements/1.1/"
_NS_DCTERMS = "http://purl.org/dc/terms/"
_NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
_NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

_ILLEGAL_SHEET_CHARS = re.compile(r"[\[\]\*?:/\\]")
_ILLEGAL_XML = re.compile(
    "[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]"
)
_RESERVED_SHEET_NAMES = {"history"}


def clean_text(value: object) -> str:
    """XML 1.0 で書けない制御文字を除き、改行を LF に揃える。"""
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return _ILLEGAL_XML.sub("", text)


def xml_text(value: object) -> str:
    """セル本文用。改行は Excel が折り返す &#10; にする。"""
    text = clean_text(value)
    escaped = _xml_escape(text)
    return escaped.replace("\n", "&#10;")


def xml_attr(value: object) -> str:
    return _xml_escape(clean_text(value), {"\n": " "})


def display_width(text: str) -> int:
    """Meiryo UI でのおおよその表示幅。全角は 2、半角は 1。"""
    width = 0
    for ch in text:
        code = ord(ch)
        if code < 0x20:
            continue
        if (
            0x1100 <= code <= 0x115F
            or 0x2329 <= code <= 0x232A
            or 0x2E80 <= code <= 0xA4CF
            or 0xAC00 <= code <= 0xD7A3
            or 0xF900 <= code <= 0xFAFF
            or 0xFE10 <= code <= 0xFE19
            or 0xFE30 <= code <= 0xFE6F
            or 0xFF00 <= code <= 0xFF60
            or 0xFFE0 <= code <= 0xFFE6
            or 0x1F300 <= code <= 0x1F64F
            or 0x20000 <= code <= 0x3FFFD
        ):
            width += 2
        else:
            width += 1
    return width


def column_letter(index: int) -> str:
    """1 始まりの列番号を A, B, ..., AA にする。"""
    if index < 1:
        raise ValueError("column index must be >= 1")
    letters = []
    while index:
        index, rem = divmod(index - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def sanitize_sheet_name(group_name: str) -> str:
    """ロググループ名を Excel のシート名に使える文字列へ近づける。"""
    text = clean_text(group_name).strip()
    if text.startswith("/"):
        text = text[1:]
    text = _ILLEGAL_SHEET_CHARS.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip().strip("'")
    if not text:
        text = "log-group"
    if text.lower() in _RESERVED_SHEET_NAMES:
        text = text + "_log"
    return text


def unique_sheet_name(preferred: str, used: set[str]) -> str:
    """31 文字以内で、ブック内で一意なシート名を返す。used を更新する。"""
    base = sanitize_sheet_name(preferred)
    if len(base) > 31:
        suffix = "~" + hashlib.sha1(preferred.encode("utf-8")).hexdigest()[:4]
        base = base[: 31 - len(suffix)] + suffix
    candidate = base[:31]
    number = 2
    while not candidate or candidate.lower() in used:
        suffix = f"~{number}"
        candidate = base[: 31 - len(suffix)] + suffix
        number += 1
        if number > 10000:
            raise RuntimeError("シート名を一意にできません: " + preferred)
    used.add(candidate.lower())
    return candidate


def phrase_color(index: int) -> str:
    return HIGHLIGHT_COLORS[index % len(HIGHLIGHT_COLORS)]


@dataclass
class Run:
    text: str
    bold: bool = False
    color: Optional[str] = None
    size: int = 10


def highlight_runs(
    message: str,
    phrases: Sequence[str],
    pattern: Optional[re.Pattern[str]],
    ignore_case: bool,
) -> list[Run]:
    """指定文言だけを太字・別色にしたラン列を返す。一致がなければ本文 1 本。"""
    text = clean_text(message)
    if not text or pattern is None or not phrases:
        return [Run(text)]

    colors: dict[str, str] = {}
    for index, phrase in enumerate(phrases):
        key = phrase.casefold() if ignore_case else phrase
        colors.setdefault(key, phrase_color(index))

    runs: list[Run] = []
    cursor = 0
    for match in pattern.finditer(text):
        start, end = match.span()
        if start < cursor:
            continue
        if start > cursor:
            runs.append(Run(text[cursor:start]))
        matched = match.group(0)
        key = matched.casefold() if ignore_case else matched
        color = colors.get(key, phrase_color(0))
        runs.append(Run(matched, bold=True, color=color))
        cursor = end
    if cursor < len(text):
        runs.append(Run(text[cursor:]))
    return runs or [Run("")]


def compile_highlights(
    phrases: Sequence[str], ignore_case: bool
) -> Optional[re.Pattern[str]]:
    usable = [phrase for phrase in phrases if phrase]
    if not usable:
        return None
    ordered = sorted(set(usable), key=len, reverse=True)
    flags = re.MULTILINE
    if ignore_case:
        flags |= re.IGNORECASE
    return re.compile("|".join(re.escape(phrase) for phrase in ordered), flags)


def slice_runs(runs: Sequence[Run], limit: int = EXCEL_MAX_CELL_CHARS) -> list[list[Run]]:
    """ラン列をセル上限以下の行に切る。色は切断面をまたいでも維持する。"""
    if limit < 1:
        raise ValueError("limit must be positive")
    slices: list[list[Run]] = []
    current: list[Run] = []
    remaining = limit
    for run in runs:
        text = run.text
        if text == "":
            continue
        while text:
            if remaining == 0:
                slices.append(current)
                current = []
                remaining = limit
            take = text[:remaining]
            text = text[len(take) :]
            current.append(Run(take, run.bold, run.color, run.size))
            remaining -= len(take)
    if current:
        slices.append(current)
    return slices or [[Run("")]]


@dataclass
class WorkbookMeta:
    account_id: str
    caller_arn: str
    region: str
    start_display: str
    end_display: str
    timezone_name: str
    output_path: str
    generated_display: str
    generated_iso: str
    scanned_groups: int
    empty_groups: int
    ignore_case: bool
    phrases: tuple[str, ...]
    notice: str = ""


@dataclass
class StreamSource:
    name: str
    event_count: int
    first_ms: Optional[int]
    last_ms: Optional[int]
    iter_events: Callable[[], Iterable[tuple]]


def _empty_stream_iter() -> Iterable[StreamSource]:
    return ()


@dataclass
class GroupSource:
    name: str
    retention_days: Optional[int]
    stream_count: int
    event_count: int
    first_ms: Optional[int]
    last_ms: Optional[int]
    include_sheet: bool
    note: str = ""
    iter_streams: Callable[[], Iterable[StreamSource]] = _empty_stream_iter


@dataclass
class SheetResult:
    log_group: str
    sheet_name: str
    streams: int
    events: int
    first_display: str
    last_display: str
    status: str


@dataclass
class _Font:
    size: int
    color: str
    bold: bool = False
    underline: bool = False


@dataclass
class _Fill:
    rgb: Optional[str]


@dataclass
class _Xf:
    font: int
    fill: int
    border: int
    horizontal: str = "left"
    vertical: str = "center"
    wrap: bool = False
    num_fmt: int = 0


class _Styles:
    """セルスタイルを名前で引けるように積む。fills の 0 と 1 は OOXML の予約。"""

    def __init__(self) -> None:
        self.fonts: list[_Font] = []
        self.fills: list[_Fill] = [_Fill(None), _Fill("gray125")]
        self.xfs: list[_Xf] = []
        self.names: dict[str, int] = {}
        self._font_index: dict[tuple, int] = {}
        self._fill_index: dict[str, int] = {}

    def font(self, size: int, color: str, bold: bool = False, underline: bool = False) -> int:
        key = (size, color.upper(), bold, underline)
        found = self._font_index.get(key)
        if found is not None:
            return found
        self.fonts.append(_Font(size, color.upper(), bold, underline))
        index = len(self.fonts) - 1
        self._font_index[key] = index
        return index

    def fill(self, rgb: str) -> int:
        key = rgb.upper()
        found = self._fill_index.get(key)
        if found is not None:
            return found
        self.fills.append(_Fill(key))
        index = len(self.fills) - 1
        self._fill_index[key] = index
        return index

    def add(
        self,
        name: str,
        *,
        size: int,
        color: str,
        fill: str,
        bold: bool = False,
        underline: bool = False,
        border: int = 0,
        horizontal: str = "left",
        vertical: str = "center",
        wrap: bool = False,
        num_fmt: int = 0,
    ) -> int:
        index = len(self.xfs)
        self.xfs.append(
            _Xf(
                font=self.font(size, color, bold, underline),
                fill=self.fill(fill),
                border=border,
                horizontal=horizontal,
                vertical=vertical,
                wrap=wrap,
                num_fmt=num_fmt,
            )
        )
        self.names[name] = index
        return index

    def sid(self, name: str) -> int:
        return self.names[name]

    def xml(self) -> str:
        font_xml = []
        for font in self.fonts:
            parts = ["<font>"]
            if font.bold:
                parts.append("<b/>")
            if font.underline:
                parts.append("<u/>")
            parts.append(f'<sz val="{font.size}"/>')
            parts.append(f'<color rgb="FF{font.color}"/>')
            parts.append(f'<name val="{FONT_NAME}"/>')
            parts.append("<family val=\"3\"/>")
            parts.append("<charset val=\"128\"/>")
            parts.append("</font>")
            font_xml.append("".join(parts))

        fill_xml = []
        for index, fill in enumerate(self.fills):
            if index == 0:
                fill_xml.append('<fill><patternFill patternType="none"/></fill>')
            elif fill.rgb == "gray125":
                fill_xml.append('<fill><patternFill patternType="gray125"/></fill>')
            else:
                fill_xml.append(
                    "<fill><patternFill patternType=\"solid\">"
                    f'<fgColor rgb="FF{fill.rgb}"/>'
                    '<bgColor indexed="64"/>'
                    "</patternFill></fill>"
                )

        borders = [
            (
                "<border>"
                '<left/><right/><top/><bottom/><diagonal/>'
                "</border>"
            ),
            (
                "<border>"
                '<left style="thin"><color rgb="FFD5DDE3"/></left>'
                '<right style="thin"><color rgb="FFD5DDE3"/></right>'
                '<top style="thin"><color rgb="FFD5DDE3"/></top>'
                '<bottom style="thin"><color rgb="FFD5DDE3"/></bottom>'
                "<diagonal/>"
                "</border>"
            ),
        ]

        xf_xml = []
        for xf in self.xfs:
            apply_num = ' applyNumberFormat="1"' if xf.num_fmt else ""
            wrap = ' wrapText="1"' if xf.wrap else ""
            xf_xml.append(
                f'<xf numFmtId="{xf.num_fmt}" fontId="{xf.font}" fillId="{xf.fill}" '
                f'borderId="{xf.border}" xfId="0" applyFont="1" applyFill="1" '
                f'applyBorder="1" applyAlignment="1"{apply_num}>'
                f'<alignment horizontal="{xf.horizontal}" vertical="{xf.vertical}"{wrap}/>'
                "</xf>"
            )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<styleSheet xmlns="{_NS_MAIN}">'
            f'<fonts count="{len(self.fonts)}">{"".join(font_xml)}</fonts>'
            f'<fills count="{len(self.fills)}">{"".join(fill_xml)}</fills>'
            f'<borders count="{len(borders)}">{"".join(borders)}</borders>'
            '<cellStyleXfs count="1">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
            "</cellStyleXfs>"
            f'<cellXfs count="{len(self.xfs)}">{"".join(xf_xml)}</cellXfs>'
            '<cellStyles count="1">'
            '<cellStyle name="Normal" xfId="0" builtinId="0"/>'
            "</cellStyles>"
            "</styleSheet>"
        )


def build_styles() -> _Styles:
    styles = _Styles()
    navy = "1B3A4B"
    navy2 = "245066"
    teal = "0E6B68"
    teal2 = "127E7A"
    ink = "1F2933"
    muted = "5E6A75"
    white = "FFFFFF"
    wash = "E7EEF2"
    zebra = "F4F7F8"
    paper = "FFFFFF"
    section = "F7F9FA"
    header = "E4EEF1"
    link = "0C6E8A"
    thin = 1

    # 0 番は Normal 相当にしておく。
    styles.add("normal", size=10, color=ink, fill=paper)
    styles.add("title", size=22, color=white, fill=navy, bold=True, vertical="center")
    styles.add("subtitle", size=11, color="E7EEF2", fill=navy2)
    styles.add("section", size=13, color=navy, fill=wash, bold=True)
    styles.add(
        "meta_label", size=10, color="3E5160", fill=section, bold=True, border=thin
    )
    styles.add(
        "meta_value",
        size=10,
        color=ink,
        fill=paper,
        border=thin,
        wrap=True,
        vertical="center",
    )
    styles.add(
        "toc_header",
        size=10,
        color=white,
        fill=teal,
        bold=True,
        border=thin,
        horizontal="center",
    )
    styles.add("toc", size=10, color=ink, fill=paper, border=thin, wrap=True, vertical="center")
    styles.add("toc_alt", size=10, color=ink, fill="F3F8F7", border=thin, wrap=True, vertical="center")
    styles.add(
        "toc_link",
        size=10,
        color=link,
        fill=paper,
        border=thin,
        underline=True,
        wrap=True,
    )
    styles.add(
        "toc_link_alt",
        size=10,
        color=link,
        fill="F3F8F7",
        border=thin,
        underline=True,
        wrap=True,
    )
    styles.add(
        "toc_num",
        size=10,
        color=ink,
        fill=paper,
        border=thin,
        horizontal="right",
        num_fmt=3,
    )
    styles.add(
        "toc_num_alt",
        size=10,
        color=ink,
        fill="F3F8F7",
        border=thin,
        horizontal="right",
        num_fmt=3,
    )
    styles.add("group_banner", size=16, color=white, fill=navy, bold=True)
    styles.add("group_sub", size=10, color="E7EEF2", fill=navy2, wrap=True)
    styles.add("stream_banner", size=12, color=white, fill=teal, bold=True, wrap=True)
    styles.add("stream_sub", size=9, color="E7F4F3", fill=teal2, wrap=True)
    styles.add(
        "col_header",
        size=10,
        color=navy,
        fill=header,
        bold=True,
        border=thin,
        horizontal="center",
    )
    styles.add(
        "num", size=10, color=ink, fill=paper, border=thin, horizontal="right", vertical="top"
    )
    styles.add(
        "num_alt",
        size=10,
        color=ink,
        fill=zebra,
        border=thin,
        horizontal="right",
        vertical="top",
    )
    styles.add("time", size=10, color=ink, fill=paper, border=thin, vertical="top")
    styles.add("time_alt", size=10, color=ink, fill=zebra, border=thin, vertical="top")
    styles.add(
        "msg", size=10, color=ink, fill=paper, border=thin, wrap=True, vertical="top"
    )
    styles.add(
        "msg_alt", size=10, color=ink, fill=zebra, border=thin, wrap=True, vertical="top"
    )
    styles.add("eid", size=9, color=muted, fill=paper, border=thin, vertical="top")
    styles.add("eid_alt", size=9, color=muted, fill=zebra, border=thin, vertical="top")
    styles.add("note", size=10, color=muted, fill="F7F5F2", border=thin, wrap=True)
    styles.add("cover_note", size=9, color=muted, fill=paper, wrap=True, vertical="top")
    styles.add("spacer", size=10, color=white, fill=paper)
    return styles


class _Cell:
    __slots__ = ("style", "text", "runs", "number", "link")

    def __init__(
        self,
        style: int,
        text: Optional[str] = None,
        runs: Optional[Sequence[Run]] = None,
        number: Optional[int] = None,
        link: Optional[str] = None,
    ) -> None:
        self.style = style
        self.text = text
        self.runs = runs
        self.number = number
        self.link = link


class _Sheet:
    def __init__(
        self,
        directory: Path,
        sheet_id: int,
        name: str,
        tab_rgb: str,
        widths: Sequence[float],
        freeze_row: int,
    ) -> None:
        self.sheet_id = sheet_id
        self.name = name
        self.tab_rgb = tab_rgb
        self.widths = list(widths)
        self.freeze_row = freeze_row
        self.row = 0
        self.merges: list[str] = []
        self.links: list[tuple[str, str]] = []
        self.body_path = directory / f"sheet{sheet_id}.body.xml"
        self._body = self.body_path.open("w", encoding="utf-8", newline="\n")
        self._ncols = len(widths)

    def close_body(self) -> None:
        if not self._body.closed:
            self._body.close()

    def add_row(self, height: float, cells: Sequence[_Cell], merge: Optional[str] = None) -> int:
        self.row += 1
        if len(cells) != self._ncols:
            raise ValueError(
                f"{self.name} row {self.row}: expected {self._ncols} cells, got {len(cells)}"
            )
        spans = f"1:{self._ncols}"
        parts = [
            f'<row r="{self.row}" ht="{height:.1f}" customHeight="1" spans="{spans}">'
        ]
        for index, cell in enumerate(cells, start=1):
            ref = f"{column_letter(index)}{self.row}"
            if cell.link:
                self.links.append((ref, cell.link))
            if cell.number is not None:
                parts.append(f'<c r="{ref}" s="{cell.style}"><v>{cell.number}</v></c>')
            elif cell.runs is not None:
                parts.append(
                    f'<c r="{ref}" s="{cell.style}" t="inlineStr"><is>'
                    f"{_runs_xml(cell.runs)}</is></c>"
                )
            elif cell.text is not None:
                parts.append(
                    f'<c r="{ref}" s="{cell.style}" t="inlineStr"><is>'
                    f'<t xml:space="preserve">{xml_text(cell.text)}</t></is></c>'
                )
            else:
                parts.append(f'<c r="{ref}" s="{cell.style}"/>')
        parts.append("</row>")
        self._body.write("".join(parts))
        self._body.write("\n")
        if merge:
            self.merges.append(merge.replace("#", str(self.row)))
        return self.row

    def xml_around_body(self) -> tuple[str, str]:
        cols = []
        for index, width in enumerate(self.widths, start=1):
            cols.append(
                f'<col min="{index}" max="{index}" width="{width:.2f}" customWidth="1"/>'
            )
        last_col = column_letter(self._ncols)
        last_row = max(1, self.row)
        if self.freeze_row > 0:
            selected = "1" if self.sheet_id == 1 else "0"
            view = (
                f'<sheetViews><sheetView workbookViewId="0" showGridLines="0" '
                f'tabSelected="{selected}">'
                f'<pane ySplit="{self.freeze_row}" topLeftCell="A{self.freeze_row + 1}" '
                f'activePane="bottomLeft" state="frozen"/>'
                f'<selection pane="bottomLeft" activeCell="A{self.freeze_row + 1}" '
                f'sqref="A{self.freeze_row + 1}"/>'
                "</sheetView></sheetViews>"
            )
        else:
            selected = "1" if self.sheet_id == 1 else "0"
            view = (
                f'<sheetViews><sheetView workbookViewId="0" showGridLines="0" '
                f'tabSelected="{selected}"/></sheetViews>'
            )
        header = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
            "<sheetPr>"
            f'<tabColor rgb="FF{self.tab_rgb}"/>'
            '<pageSetUpPr fitToPage="1"/>'
            "</sheetPr>"
            f'<dimension ref="A1:{last_col}{last_row}"/>'
            f"{view}"
            '<sheetFormatPr defaultRowHeight="18" customHeight="1"/>'
            f'<cols>{"".join(cols)}</cols>'
            "<sheetData>\n"
        )
        footer_parts = ["</sheetData>"]
        if self.merges:
            footer_parts.append(f'<mergeCells count="{len(self.merges)}">')
            for ref in self.merges:
                footer_parts.append(f'<mergeCell ref="{xml_attr(ref)}"/>')
            footer_parts.append("</mergeCells>")
        if self.links:
            footer_parts.append("<hyperlinks>")
            for ref, location in self.links:
                footer_parts.append(
                    f'<hyperlink ref="{xml_attr(ref)}" location="{xml_attr(location)}"/>'
                )
            footer_parts.append("</hyperlinks>")
        footer_parts.append(
            '<pageMargins left="0.45" right="0.45" top="0.7" bottom="0.6" '
            'header="0.28" footer="0.28"/>'
        )
        footer_parts.append(
            '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" '
            'fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>'
        )
        footer_parts.append(
            "<headerFooter>"
            "<oddHeader>&amp;LCloudWatch Logs&amp;C&amp;A&amp;R&amp;D &amp;T</oddHeader>"
            "<oddFooter>&amp;LCloudWatch Logs Export&amp;C&amp;P / &amp;N&amp;R&amp;A</oddFooter>"
            "</headerFooter>"
            "</worksheet>"
        )
        return header, "".join(footer_parts)


def _runs_xml(runs: Sequence[Run]) -> str:
    parts = []
    for run in runs:
        rpr = ["<rPr>"]
        if run.bold:
            rpr.append("<b/>")
        rpr.append(f'<sz val="{run.size}"/>')
        color = run.color or "1F2933"
        rpr.append(f'<color rgb="FF{color}"/>')
        rpr.append(f'<rFont val="{FONT_NAME}"/>')
        rpr.append("<family val=\"3\"/>")
        rpr.append("<charset val=\"128\"/>")
        rpr.append("</rPr>")
        parts.append(
            f'<r>{"".join(rpr)}<t xml:space="preserve">{xml_text(run.text)}</t></r>'
        )
    if not parts:
        parts.append(
            f'<r><rPr><sz val="10"/><color rgb="FF1F2933"/>'
            f'<rFont val="{FONT_NAME}"/><charset val="128"/></rPr>'
            f'<t xml:space="preserve"></t></r>'
        )
    return "".join(parts)


def _filled(style: int, count: int, text: Optional[str] = None, runs=None) -> list[_Cell]:
    cells = [_Cell(style, text=text, runs=runs)]
    for _ in range(count - 1):
        cells.append(_Cell(style))
    return cells


def _merge_ref(start_col: int, end_col: int) -> str:
    return f"{column_letter(start_col)}#:{column_letter(end_col)}#"


def _sheet_link(name: str) -> str:
    return "'" + name.replace("'", "''") + "'!A1"


def _row_height_for_message(message: str, col_width: float) -> float:
    usable = max(12, int(col_width) - 1)
    lines = 0
    for line in clean_text(message).split("\n"):
        width = max(1, display_width(line))
        lines += max(1, (width + usable - 1) // usable)
    return float(min(409, max(18, lines * 15)))


def _message_column_width(group: GroupSource) -> float:
    """ログが見える幅。極端に長い 1 行でシート全体を広げない。"""
    if group.event_count <= 0:
        return 88.0
    sample_cap = 100000 if group.event_count <= 100000 else 5000
    widths: list[int] = []
    seen = 0
    for stream in group.iter_streams():
        for _ts, message, _event_id in stream.iter_events():
            for line in clean_text(message).split("\n"):
                widths.append(max(1, display_width(line)))
            seen += 1
            if seen >= sample_cap:
                break
        if seen >= sample_cap:
            break
    if not widths:
        return 88.0
    widths.sort()
    index = min(len(widths) - 1, max(0, int(round((len(widths) - 1) * 0.90))))
    chosen = widths[index] + 3
    return float(max(48, min(120, chosen)))


def _retention_label(days: Optional[int]) -> str:
    if days is None:
        return "無期限"
    return f"{days} 日"


def write_xlsx(
    path: Path,
    *,
    meta: WorkbookMeta,
    groups: Sequence[GroupSource],
    format_ts: Callable[[int], str],
    max_sheet_rows: int = EXCEL_MAX_ROW,
) -> list[SheetResult]:
    """グループ一覧を受け取り、xlsx を path に書く。戻り値はシートごとの件数。"""
    if max_sheet_rows < 8 or max_sheet_rows > EXCEL_MAX_ROW:
        raise ValueError("max_sheet_rows が Excel の範囲外です")
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_suffix(path.suffix + ".partial")
    styles = build_styles()
    pattern = compile_highlights(meta.phrases, meta.ignore_case)
    results: list[SheetResult] = []
    used_names: set[str] = set()
    work_dir = Path(path.parent) / (".cwlogs-" + hashlib.sha1(os.urandom(8)).hexdigest()[:10])
    work_dir.mkdir(mode=0o700)
    sheets: list[_Sheet] = []
    try:
        cover = _Sheet(
            work_dir,
            sheet_id=1,
            name=unique_sheet_name("概要", used_names),
            tab_rgb="1B3A4B",
            widths=(46, 28, 16, 14, 24, 24, 28),
            freeze_row=0,
        )
        sheets.append(cover)
        for group in groups:
            if not group.include_sheet:
                continue
            _write_group_sheets(
                work_dir=work_dir,
                sheets=sheets,
                group=group,
                used_names=used_names,
                styles=styles,
                pattern=pattern,
                meta=meta,
                format_ts=format_ts,
                max_sheet_rows=max_sheet_rows,
                results=results,
            )
        _write_cover(cover, styles, meta, results)
        cover.close_body()
        _package(path, sheets, styles, meta)
    except Exception:
        if partial.exists():
            partial.unlink()
        raise
    finally:
        for sheet in sheets:
            try:
                sheet.close_body()
            except Exception:
                pass
        shutil.rmtree(work_dir, ignore_errors=True)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    return results


def _write_cover(
    cover: _Sheet,
    styles: _Styles,
    meta: WorkbookMeta,
    results: Sequence[SheetResult],
) -> None:
    sid = styles.sid
    ncols = 7
    title = _filled(sid("title"), ncols, "CloudWatch Logs エクスポート")
    cover.add_row(36, title, _merge_ref(1, ncols))
    subtitle = _filled(
        sid("subtitle"),
        ncols,
        "指定した時間範囲に出力されたロググループとログストリーム",
    )
    cover.add_row(22, subtitle, _merge_ref(1, ncols))
    cover.add_row(10, _filled(sid("spacer"), ncols))

    section = _filled(sid("section"), ncols, "抽出条件")
    cover.add_row(22, section, _merge_ref(1, ncols))

    exported_events = sum(item.events for item in results)
    exported_streams = sum(item.streams for item in results)
    phrase_cell = _phrase_legend(meta)
    rows = [
        ("アカウント ID", meta.account_id or "（不明）"),
        ("プリンシパル", meta.caller_arn or "（不明）"),
        ("リージョン", meta.region),
        ("開始", meta.start_display + "  から（この時刻を含む）"),
        ("終了", meta.end_display + "  まで（この時刻を含む）"),
        ("表示タイムゾーン", meta.timezone_name),
        ("生成日時", meta.generated_display),
        ("出力ファイル", meta.output_path),
        ("走査したロググループ", str(meta.scanned_groups)),
        ("期間内イベントなし", str(meta.empty_groups)),
        ("出力イベント数", f"{exported_events:,}"),
        ("出力ストリーム数", f"{exported_streams:,}"),
        ("大文字小文字", "区別しない" if meta.ignore_case else "区別する"),
    ]
    if meta.notice:
        rows.append(("注意", meta.notice))
    for label, value in rows:
        cells = [
            _Cell(sid("meta_label"), text=label),
            _Cell(sid("meta_value"), text=value),
            _Cell(sid("meta_value")),
            _Cell(sid("meta_value")),
            _Cell(sid("meta_value")),
            _Cell(sid("meta_value")),
            _Cell(sid("meta_value")),
        ]
        cover.add_row(20 if len(value) < 80 else 32, cells, _merge_ref(2, ncols))
    legend_cells = [
        _Cell(sid("meta_label"), text="強調する文言"),
        _Cell(sid("meta_value"), runs=phrase_cell),
        _Cell(sid("meta_value")),
        _Cell(sid("meta_value")),
        _Cell(sid("meta_value")),
        _Cell(sid("meta_value")),
        _Cell(sid("meta_value")),
    ]
    legend_height = 20 if len(meta.phrases) <= 3 else min(80, 18 + 14 * len(meta.phrases))
    cover.add_row(legend_height, legend_cells, _merge_ref(2, ncols))

    cover.add_row(10, _filled(sid("spacer"), ncols))
    cover.add_row(22, _filled(sid("section"), ncols, "ロググループ一覧"), _merge_ref(1, ncols))
    headers = ("ロググループ", "シート", "ストリーム", "イベント", "先頭", "末尾", "状態")
    header_cells = [_Cell(sid("toc_header"), text=text) for text in headers]
    header_row = cover.add_row(20, header_cells)
    cover.freeze_row = header_row

    if not results:
        empty = _filled(
            sid("note"),
            ncols,
            "指定した時間範囲に出力されたログはありません。",
        )
        cover.add_row(24, empty, _merge_ref(1, ncols))
    for index, item in enumerate(results):
        alt = index % 2 == 1
        link_style = sid("toc_link_alt" if alt else "toc_link")
        text_style = sid("toc_alt" if alt else "toc")
        num_style = sid("toc_num_alt" if alt else "toc_num")
        height = 20
        if display_width(item.log_group) > 42 or len(item.status) > 18:
            height = 32
        cells = [
            _Cell(text_style, text=item.log_group),
            _Cell(link_style, text=item.sheet_name, link=_sheet_link(item.sheet_name)),
            _Cell(num_style, number=item.streams),
            _Cell(num_style, number=item.events),
            _Cell(text_style, text=item.first_display),
            _Cell(text_style, text=item.last_display),
            _Cell(text_style, text=item.status),
        ]
        cover.add_row(height, cells)

    cover.add_row(12, _filled(sid("spacer"), ncols))
    note = (
        "各ロググループは別シートです。ログストリームはシート内の色帯ブロックです。"
        "強調した文言だけが太字と別色になり、ログ本文そのものは改変していません。"
        "1 セルに収まらない長い行は、色を保ったまま次の行へ分割しています。"
        "フォントは Meiryo UI です。開く PC にこのフォントがない場合は、Excel が近いフォントで置き換えます。"
    )
    cover.add_row(48, _filled(sid("cover_note"), ncols, note), _merge_ref(1, ncols))


def _phrase_legend(meta: WorkbookMeta) -> list[Run]:
    if not meta.phrases:
        return [Run("（指定なし）", size=10)]
    runs: list[Run] = []
    for index, phrase in enumerate(meta.phrases):
        if index:
            runs.append(Run("    ", size=10))
        runs.append(Run(phrase, bold=True, color=phrase_color(index), size=10))
    runs.append(Run("    ログ本文中の一致箇所だけに適用", size=10, color="5E6A75"))
    return runs


def _write_group_sheets(
    *,
    work_dir: Path,
    sheets: list[_Sheet],
    group: GroupSource,
    used_names: set[str],
    styles: _Styles,
    pattern: Optional[re.Pattern[str]],
    meta: WorkbookMeta,
    format_ts: Callable[[int], str],
    max_sheet_rows: int,
    results: list[SheetResult],
) -> list[str]:
    sid = styles.sid
    message_width = _message_column_width(group)
    widths = (8.5, 24.0, message_width, 36.0)
    names: list[str] = []
    state = {
        "sheet": None,
        "part": 0,
        "events": 0,
        "streams": 0,
        "seen_streams": set(),
    }

    def note_stream(name: str) -> None:
        seen = state["seen_streams"]
        if name not in seen:
            seen.add(name)
            state["streams"] = int(state["streams"]) + 1

    def first_last() -> tuple[str, str]:
        first = format_ts(group.first_ms) if group.first_ms is not None else ""
        last = format_ts(group.last_ms) if group.last_ms is not None else ""
        return first, last

    def open_sheet(continued: bool) -> _Sheet:
        state["part"] += 1
        preferred = group.name if state["part"] == 1 else f"{group.name}~{state['part']}"
        name = unique_sheet_name(preferred, used_names)
        sheet = _Sheet(
            work_dir,
            sheet_id=len(sheets) + 1,
            name=name,
            tab_rgb="0E6B68",
            widths=widths,
            freeze_row=2,
        )
        sheets.append(sheet)
        names.append(name)
        banner = group.name if not continued else group.name + "    （続き）"
        sheet.add_row(30, _filled(sid("group_banner"), 4, "ロググループ   " + banner), _merge_ref(1, 4))
        sub = (
            f"リージョン {meta.region}    "
            f"期間 {meta.start_display}  〜  {meta.end_display} ({meta.timezone_name})    "
            f"保持 {_retention_label(group.retention_days)}    "
            f"ストリーム {group.stream_count:,}    イベント {group.event_count:,}"
        )
        sheet.add_row(20, _filled(sid("group_sub"), 4, sub), _merge_ref(1, 4))
        first, last = first_last()
        if continued:
            status = "続き"
        elif group.note and group.event_count <= 0:
            status = group.note
        elif group.note:
            status = group.note
        else:
            status = "出力"
        state["events"] = 0
        state["streams"] = 0
        state["seen_streams"] = set()
        results.append(
            SheetResult(
                log_group=group.name,
                sheet_name=name,
                streams=0,
                events=0,
                first_display=first,
                last_display=last,
                status=status,
            )
        )
        state["sheet"] = sheet
        return sheet

    def current() -> _Sheet:
        sheet = state["sheet"]
        if sheet is None:
            return open_sheet(False)
        return sheet

    def room(needed: int) -> _Sheet:
        sheet = current()
        if sheet.row + needed <= max_sheet_rows:
            return sheet
        sheet.close_body()
        _update_result_counts(results, sheet.name, state)
        return open_sheet(True)

    if group.note and group.event_count <= 0:
        sheet = open_sheet(False)
        sheet.add_row(10, _filled(sid("spacer"), 4))
        sheet.add_row(28, _filled(sid("note"), 4, group.note), _merge_ref(1, 4))
        sheet.close_body()
        return names

    wrote_stream = False
    for stream in group.iter_streams():
        wrote_stream = True
        sheet = room(4)
        note_stream(stream.name)
        _write_stream_header(sheet, sid, stream, format_ts, meta.timezone_name)
        number = 0
        for ts, message, event_id in stream.iter_events():
            number += 1
            runs = highlight_runs(message, meta.phrases, pattern, meta.ignore_case)
            pieces = slice_runs(runs)
            for piece_index, piece in enumerate(pieces):
                if sheet.row + 1 > max_sheet_rows:
                    sheet.close_body()
                    _update_result_counts(results, sheet.name, state)
                    sheet = open_sheet(True)
                    note_stream(stream.name)
                    _write_stream_header(
                        sheet,
                        sid,
                        stream,
                        format_ts,
                        meta.timezone_name,
                        continued=True,
                    )
                piece_text = "".join(run.text for run in piece)
                height = _row_height_for_message(piece_text, message_width)
                alt = number % 2 == 0
                no_label = str(number) if piece_index == 0 else str(number)
                eid = event_id if piece_index == 0 else "（続き）"
                cells = [
                    _Cell(sid("num_alt" if alt else "num"), text=no_label),
                    _Cell(
                        sid("time_alt" if alt else "time"),
                        text=format_ts(int(ts)),
                    ),
                    _Cell(sid("msg_alt" if alt else "msg"), runs=piece),
                    _Cell(sid("eid_alt" if alt else "eid"), text=eid),
                ]
                sheet.add_row(height, cells)
                state["events"] += 1
        if sheet.row < max_sheet_rows:
            sheet.add_row(8, _filled(sid("spacer"), 4))

    if not wrote_stream:
        sheet = open_sheet(False)
        sheet.add_row(10, _filled(sid("spacer"), 4))
        note = group.note or "この期間に出力されたイベントはありません。"
        sheet.add_row(28, _filled(sid("note"), 4, note), _merge_ref(1, 4))

    sheet = state["sheet"]
    if sheet is not None and not sheet._body.closed:
        sheet.close_body()
        _update_result_counts(results, sheet.name, state)
    return names


def _update_result_counts(results: list[SheetResult], sheet_name: str, state: dict) -> None:
    for item in results:
        if item.sheet_name == sheet_name and item.events == 0 and item.streams == 0:
            item.events = int(state["events"])
            item.streams = int(state["streams"])
            state["events"] = 0
            state["streams"] = 0
            return


def _write_stream_header(
    sheet: _Sheet,
    sid: Callable[[str], int],
    stream: StreamSource,
    format_ts: Callable[[int], str],
    timezone_name: str,
    continued: bool = False,
) -> None:
    title = "ログストリーム   " + stream.name
    if continued:
        title += "    （続き）"
    height = 24
    if display_width(stream.name) > 90:
        height = 36
    sheet.add_row(height, _filled(sid("stream_banner"), 4, title), _merge_ref(1, 4))
    first = format_ts(stream.first_ms) if stream.first_ms is not None else "—"
    last = format_ts(stream.last_ms) if stream.last_ms is not None else "—"
    sub = (
        f"件数 {stream.event_count:,}    "
        f"先頭 {first}    末尾 {last}    時刻は {timezone_name}"
    )
    sheet.add_row(18, _filled(sid("stream_sub"), 4, sub), _merge_ref(1, 4))
    headers = (
        "No.",
        f"タイムスタンプ ({timezone_name})",
        "メッセージ",
        "イベント ID",
    )
    sheet.add_row(20, [_Cell(sid("col_header"), text=text) for text in headers])


def _package(path: Path, sheets: Sequence[_Sheet], styles: _Styles, meta: WorkbookMeta) -> None:
    sheet_entries = []
    workbook_sheets = []
    workbook_rels = []
    overrides = []
    for index, sheet in enumerate(sheets, start=1):
        rid = f"rId{index}"
        target = f"worksheets/sheet{index}.xml"
        sheet_entries.append((f"xl/{target}", sheet))
        workbook_sheets.append(
            f'<sheet name="{xml_attr(sheet.name)}" sheetId="{index}" r:id="{rid}"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="{rid}" '
            f'Type="{_NS_REL}/worksheet" Target="{target}"/>'
        )
        overrides.append(
            f'<Override PartName="/xl/{target}" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    style_rid = f"rId{len(sheets) + 1}"
    workbook_rels.append(
        f'<Relationship Id="{style_rid}" Type="{_NS_REL}/styles" Target="styles.xml"/>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
        '<workbookPr date1904="0"/>'
        '<bookViews><workbookView activeTab="0" windowWidth="28000" windowHeight="15000"/></bookViews>'
        f'<sheets>{"".join(workbook_sheets)}</sheets>'
        "</workbook>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG}">'
        f"{''.join(workbook_rels)}"
        "</Relationships>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG}">'
        f'<Relationship Id="rId1" Type="{_NS_REL}/officeDocument" Target="xl/workbook.xml"/>'
        f'<Relationship Id="rId2" Type="{_NS_PKG}/metadata/core-properties" Target="docProps/core.xml"/>'
        f'<Relationship Id="rId3" Type="{_NS_REL}/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{_NS_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        f"{''.join(overrides)}"
        "</Types>"
    )
    created = meta.generated_iso or _iso_now_from_display(meta.generated_display)
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<cp:coreProperties xmlns:cp="{_NS_CP}" xmlns:dc="{_NS_DC}" '
        f'xmlns:dcterms="{_NS_DCTERMS}" xmlns:xsi="{_NS_XSI}">'
        "<dc:title>CloudWatch Logs エクスポート</dc:title>"
        "<dc:subject>CloudWatch Logs</dc:subject>"
        "<dc:creator>cwlogs-export</dc:creator>"
        "<cp:lastModifiedBy>cwlogs-export</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
        "</cp:coreProperties>"
    )
    titles = "".join(f"<vt:lpstr>{xml_text(sheet.name)}</vt:lpstr>" for sheet in sheets)
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{_NS_EP}" xmlns:vt="{_NS_VT}">'
        "<Application>cwlogs-export</Application>"
        "<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">"
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"
        f"<vt:variant><vt:i4>{len(sheets)}</vt:i4></vt:variant>"
        "</vt:vector></HeadingPairs>"
        f'<TitlesOfParts><vt:vector size="{len(sheets)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
        "</Properties>"
    )

    partial = path.with_suffix(path.suffix + ".partial")
    if partial.exists():
        partial.unlink()
    with zipfile.ZipFile(partial, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
        _writestr(archive, "[Content_Types].xml", content_types)
        _writestr(archive, "_rels/.rels", root_rels)
        _writestr(archive, "docProps/core.xml", core)
        _writestr(archive, "docProps/app.xml", app)
        _writestr(archive, "xl/workbook.xml", workbook_xml)
        _writestr(archive, "xl/_rels/workbook.xml.rels", rels_xml)
        _writestr(archive, "xl/styles.xml", styles.xml())
        for entry, sheet in sheet_entries:
            header, footer = sheet.xml_around_body()
            with archive.open(entry, "w") as raw:
                raw.write(header.encode("utf-8"))
                with sheet.body_path.open("rb") as body:
                    shutil.copyfileobj(body, raw, length=1024 * 1024)
                raw.write(footer.encode("utf-8"))
    os.replace(partial, path)


def _writestr(archive: zipfile.ZipFile, name: str, text: str) -> None:
    archive.writestr(name, text.encode("utf-8"))


def _iso_now_from_display(display: str) -> str:
    """core.xml 用。表示文字列から厳密に戻さない。呼び出し側が ISO を渡せないときの予備。"""
    text = display.strip()
    if re.match(r"\d{4}-\d{2}-\d{2}T", text):
        return clean_text(text)
    return "1970-01-01T00:00:00Z"


def iter_events_list(events: Sequence[tuple]) -> Callable[[], Iterator[tuple]]:
    """テストや小さな入力用。何度呼んでも同じイベントを返す。"""

    def factory() -> Iterator[tuple]:
        return iter(events)

    return factory
