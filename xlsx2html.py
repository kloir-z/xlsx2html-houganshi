#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx2html.py — Excel 方眼紙ドキュメントを「見た目そのまま」の単一 HTML に変換する。

方針:
  * 列幅 / 行高 をピクセルに換算し、table-layout:fixed の表として厳密に再現する
    （= 方眼のインデントがそのまま残る）
  * 罫線・塗り・フォント・文字揃え・結合セル・縦書き・回転を CSS で再現
  * 数式は評価済みキャッシュ値を使用（data_only=True）。表示は表示形式に従って整形
  * セル外にはみ出す文字（Excel のオーバーフロー）も再現するため、
    文字は絶対配置の <span> として置き、セル幅に縛られないようにしている
  * 画像・オートシェイプは xl/drawings/*.xml を直接解析し、
    絶対座標のオーバーレイとして重ねる（画像は base64 で埋め込み → 単一ファイル完結）
  * 文字はすべて実テキストなのでコピー可能

使い方:
  python xlsx2html.py input.xlsx [-o output.html] [--sheet 名前] [--gridlines auto|on|off]
"""

from __future__ import annotations

import argparse
import base64
import colorsys
import html
import math
import os
import posixpath
import re
import sys
import warnings
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, time, timedelta

# openpyxl が読み飛ばす拡張機能（条件付き書式の拡張・ヘッダーフッター等）の
# 警告は変換結果に影響しないので黙らせる
warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl.*")

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl が必要です:  pip install openpyxl")


# ---------------------------------------------------------------------------
# 単位換算
# ---------------------------------------------------------------------------

EMU_PER_PX = 9525.0          # 1px = 9525 EMU (96dpi)
MDW = 7                      # 標準フォントの数字1文字幅(px)。Calibri 11 / MS Pゴシック 11 とも 7
DEFAULT_COL_WIDTH = 8.43     # Excel 既定の列幅(文字数)
DEFAULT_ROW_HEIGHT_PT = 18.75
GRID_COLOR = "#e8e8e8"       # Excel の目盛線相当（画面表示のみ、既定では印刷されない）
CELL_PAD = 2                 # セルの左右パディング(px)。Excel 実測相当
INDENT_PX = 8                # インデント1段あたりの px


def col_width_to_px(width: float | None) -> int:
    """Excel の列幅(文字数)→ px。既定 8.43 が 64px になる換算。"""
    if width is None:
        width = DEFAULT_COL_WIDTH
    if width <= 0:
        return 0
    return int(round(width * MDW)) + 5


def pt_to_px(pt: float) -> float:
    return pt * 4.0 / 3.0


def emu_to_px(emu) -> float:
    try:
        return float(emu) / EMU_PER_PX
    except (TypeError, ValueError):
        return 0.0


def esc(s: str) -> str:
    return html.escape(str(s), quote=True)


# ---------------------------------------------------------------------------
# 色の解決（rgb / indexed / theme + tint）
# ---------------------------------------------------------------------------

# Excel の indexed カラー表
INDEXED_COLORS = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
]

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class ColorResolver:
    """テーマカラー・インデックスカラー・tint を #rrggbb に解決する。"""

    # theme 属性の index → theme1.xml の clrScheme 要素名
    THEME_ORDER = ["lt1", "dk1", "lt2", "dk2",
                   "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                   "hlink", "folHlink"]

    def __init__(self, theme_xml: bytes | None):
        self.theme: dict[str, str] = {}
        if theme_xml:
            try:
                root = ET.fromstring(theme_xml)
                scheme = root.find(f".//{{{A_NS}}}clrScheme")
                if scheme is not None:
                    for child in scheme:
                        name = child.tag.split("}")[-1]
                        srgb = child.find(f"{{{A_NS}}}srgbClr")
                        sysc = child.find(f"{{{A_NS}}}sysClr")
                        if srgb is not None and srgb.get("val"):
                            self.theme[name] = srgb.get("val")
                        elif sysc is not None:
                            self.theme[name] = sysc.get("lastClr") or "000000"
            except ET.ParseError:
                pass
        # 既定テーマ（Office）
        self.theme.setdefault("dk1", "000000")
        self.theme.setdefault("lt1", "FFFFFF")
        self.theme.setdefault("dk2", "44546A")
        self.theme.setdefault("lt2", "E7E6E6")
        for i, c in enumerate(["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]):
            self.theme.setdefault(f"accent{i + 1}", c)
        self.theme.setdefault("hlink", "0563C1")
        self.theme.setdefault("folHlink", "954F72")

    def theme_rgb(self, idx: int) -> str:
        try:
            return self.theme.get(self.THEME_ORDER[idx], "000000")
        except IndexError:
            return "000000"

    def resolve(self, color) -> str | None:
        """openpyxl の Color オブジェクト → '#rrggbb'。自動色/未指定は None。"""
        if color is None:
            return None
        ctype = getattr(color, "type", None)
        rgb = None
        if ctype == "rgb":
            v = color.rgb
            if isinstance(v, str) and len(v) >= 6:
                rgb = v[-6:]
        elif ctype == "indexed":
            i = color.indexed
            if isinstance(i, int):
                if i == 64 or i == 65:      # 64=自動(前景) 65=自動(背景)
                    return None
                if 0 <= i < len(INDEXED_COLORS):
                    rgb = INDEXED_COLORS[i]
        elif ctype == "theme":
            t = color.theme
            if isinstance(t, int):
                rgb = self.theme_rgb(t)
        elif ctype == "auto":
            return None
        if rgb is None:
            return None
        tint = getattr(color, "tint", 0.0) or 0.0
        if tint:
            rgb = apply_tint(rgb, tint)
        return "#" + rgb.lower()


def apply_tint(rgb: str, tint: float) -> str:
    """OOXML の tint（明度補正）を適用する。"""
    r, g, b = (int(rgb[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    l = min(1.0, max(0.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def mix(c1: str, c2: str, ratio: float) -> str:
    """#rrggbb 同士を ratio で混色（ratio=1 で c1）。"""
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join("%02x" % round(x * ratio + y * (1 - ratio)) for x, y in zip(a, b))


# ---------------------------------------------------------------------------
# 表示形式（数値・日付）
# ---------------------------------------------------------------------------

JP_WEEK = ["月", "火", "水", "木", "金", "土", "日"]
ERAS = [  # (開始日, 正式名, 略名, アルファベット)
    (date(2019, 5, 1), "令和", "令", "R"),
    (date(1989, 1, 8), "平成", "平", "H"),
    (date(1926, 12, 25), "昭和", "昭", "S"),
    (date(1912, 7, 30), "大正", "大", "T"),
    (date(1868, 1, 25), "明治", "明", "M"),
]


def _split_sections(fmt: str) -> list[str]:
    out, cur, i, n = [], "", 0, len(fmt)
    in_q = in_br = False
    while i < n:
        c = fmt[i]
        if c == "\\" and i + 1 < n:
            cur += fmt[i:i + 2]
            i += 2
            continue
        if c == '"' and not in_br:
            in_q = not in_q
        elif c == "[" and not in_q:
            in_br = True
        elif c == "]" and in_br:
            in_br = False
        elif c == ";" and not in_q and not in_br:
            out.append(cur)
            cur = ""
            i += 1
            continue
        cur += c
        i += 1
    out.append(cur)
    return out


_COND_RE = re.compile(r"\[(<=|>=|<>|<|>|=)([-+0-9.eE]+)\]")


def _pick_section(sections: list[str], value: float) -> tuple[str, bool]:
    """値に対応するセクションを返す。(書式, 符号を自前で付けるか)"""
    # 条件付きセクション ([>=100] など) を優先評価
    conditional = [(i, s) for i, s in enumerate(sections) if _COND_RE.search(s)]
    if conditional:
        for i, s in conditional:
            m = _COND_RE.search(s)
            op, num = m.group(1), float(m.group(2))
            ok = {"<": value < num, "<=": value <= num, ">": value > num,
                  ">=": value >= num, "=": value == num, "<>": value != num}[op]
            if ok:
                return s, False
        rest = [s for i, s in enumerate(sections) if not _COND_RE.search(s)]
        if rest:
            return rest[0], False
    if len(sections) == 1:
        return sections[0], value < 0
    if value < 0:
        return (sections[1], False) if len(sections) >= 2 and sections[1].strip() else (sections[0], True)
    if value == 0 and len(sections) >= 3:
        return sections[2], False
    return sections[0], False


def _strip_brackets(fmt: str) -> str:
    return re.sub(r"\[(?!h+\]|m+\]|s+\])[^\]]*\]", "", fmt, flags=re.IGNORECASE)


_MASKCHARS = set("#0?.,")


def format_number(value: float, fmt: str) -> str:
    fmt = (fmt or "General").strip()
    if fmt in ("General", "@", "") or fmt.lower() == "general":
        return _general(value)
    sections = _split_sections(fmt)
    sec, own_sign = _pick_section(sections, value)
    sec = _strip_brackets(sec)
    if not sec.strip():
        return ""
    if "E" in sec.upper() and re.search(r"[#0]E[+-]", sec, re.IGNORECASE):
        return _sci(value, sec)
    if re.search(r"[#0?]\s*/\s*[#0?]", sec):     # 分数書式は近似
        return _general(value)

    v = abs(value) if not own_sign else abs(value)
    # トークン分解
    tokens: list[tuple[str, str]] = []
    i, n = 0, len(sec)
    mask = ""
    mask_done = False
    while i < n:
        c = sec[i]
        if c == "\\" and i + 1 < n:
            tokens.append(("lit", sec[i + 1]))
            i += 2
            continue
        if c == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            tokens.append(("lit", sec[i + 1:j]))
            i = j + 1
            continue
        if c in "_*":
            tokens.append(("lit", " " if c == "_" else ""))
            i += 2
            continue
        if c == "%":
            tokens.append(("lit", "%"))
            i += 1
            continue
        if c in _MASKCHARS and not mask_done:
            j = i
            while j < n and sec[j] in _MASKCHARS:
                j += 1
            mask = sec[i:j]
            tokens.append(("num", ""))
            mask_done = True
            i = j
            continue
        if c in _MASKCHARS:      # 2つ目以降のマスクは literal 扱い
            tokens.append(("lit", c))
            i += 1
            continue
        tokens.append(("lit", c))
        i += 1

    if "%" in sec.replace('\\%', ""):
        v *= 100

    if not mask:
        return "".join(t[1] for t in tokens)

    int_mask, dot, dec_mask = mask.partition(".")
    trailing = len(int_mask) - len(int_mask.rstrip(","))
    int_mask = int_mask.rstrip(",")
    grouping = "," in int_mask
    int_mask = int_mask.replace(",", "")
    dec_mask = dec_mask.replace(",", "")
    if trailing:
        v /= 1000.0 ** trailing

    decimals = len([c for c in dec_mask if c in "0#?"])
    req_dec = len([c for c in dec_mask if c in "0?"])
    body = f"{v:.{decimals}f}"
    ip, _, dp = body.partition(".")
    if decimals and req_dec < decimals:
        dp = dp[:req_dec] + dp[req_dec:].rstrip("0")
    min_int = int_mask.count("0")
    if len(ip) < min_int:
        ip = ip.rjust(min_int, "0")
    if min_int == 0 and ip == "0" and (dp or dot):
        ip = ""
    if grouping and ip:
        neg = ip.startswith("-")
        digits = ip.lstrip("-")
        digits = "{:,}".format(int(digits)) if digits.isdigit() else digits
        ip = ("-" if neg else "") + digits
    num = ip + (("." + dp) if dp else ("." if (dot and req_dec) else ""))
    if own_sign and value < 0:
        num = "-" + num
    return "".join(num if kind == "num" else txt for kind, txt in tokens)


def _general(value) -> str:
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return str(value)
        if float(value).is_integer() and abs(value) < 1e15:
            return str(int(value))
        s = repr(round(value, 10))
        return s
    return str(value)


def _sci(value: float, fmt: str) -> str:
    m = re.search(r"E([+-])0*", fmt, re.IGNORECASE)
    dec = 2
    mm = re.match(r"[#0]*\.?([#0]*)", fmt)
    if mm:
        dec = len(mm.group(1))
    s = f"{value:.{dec}E}"
    if m and m.group(1) == "-":
        s = s.replace("E+", "E")
    return s


_DATE_TOKEN_RE = re.compile(
    r'\[h+\]|\[m+\]|\[s+\]|"[^"]*"|\\.|am/pm|a/p|ggge|gge|ge|g+|yyyy|yyyy|yy|mmmmm|mmmm|mmm|mm|m|'
    r'dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s|\.0+|e+|.', re.IGNORECASE)


def format_datetime(value, fmt: str) -> str:
    if isinstance(value, timedelta):
        base = None
        total_seconds = value.total_seconds()
        dt = None
    else:
        if isinstance(value, time):
            dt = datetime(1899, 12, 31, value.hour, value.minute, value.second, value.microsecond)
        elif isinstance(value, datetime):
            dt = value
        elif isinstance(value, date):
            dt = datetime(value.year, value.month, value.day)
        else:
            return str(value)
        total_seconds = None

    fmt = _strip_brackets_keep_elapsed(fmt or "")
    sections = _split_sections(fmt)
    fmt = sections[0] if sections else fmt
    toks = _DATE_TOKEN_RE.findall(fmt)
    low = [t.lower() for t in toks]
    ampm = any(t in ("am/pm", "a/p") for t in low)

    out = []
    for idx, tok in enumerate(toks):
        t = tok.lower()
        if t.startswith('"'):
            out.append(tok[1:-1]); continue
        if t.startswith("\\"):
            out.append(tok[1:]); continue
        if dt is None:  # 経過時間のみ
            if t.startswith("[h"):
                out.append(str(int(total_seconds // 3600))); continue
            if t.startswith("[m"):
                out.append(str(int(total_seconds // 60))); continue
            if t.startswith("[s"):
                out.append(str(int(total_seconds))); continue
            if t in ("h", "hh"):
                out.append(f"{int(total_seconds // 3600) % 24:0{len(t)}d}"); continue
            if t in ("m", "mm"):
                out.append(f"{int(total_seconds // 60) % 60:0{len(t)}d}"); continue
            if t in ("s", "ss"):
                out.append(f"{int(total_seconds) % 60:0{len(t)}d}"); continue
            out.append(tok); continue

        if t.startswith("[h"):
            out.append(str(int((dt - datetime(1899, 12, 30)).total_seconds() // 3600))); continue
        if t in ("yyyy",):
            out.append(f"{dt.year:04d}"); continue
        if t == "yy":
            out.append(f"{dt.year % 100:02d}"); continue
        if t in ("ggge", "gge", "ge", "g", "gg", "ggg"):
            era = _era(dt.date())
            out.append({"ggge": era[1], "gge": era[2], "ge": era[3],
                        "g": era[3], "gg": era[2], "ggg": era[1]}.get(t, era[1])); continue
        if t.startswith("e"):
            out.append(str(_era_year(dt.date()))); continue
        if t == "aaaa":
            out.append(JP_WEEK[dt.weekday()] + "曜日"); continue
        if t == "aaa":
            out.append(JP_WEEK[dt.weekday()]); continue
        if t == "dddd":
            out.append(dt.strftime("%A")); continue
        if t == "ddd":
            out.append(dt.strftime("%a")); continue
        if t in ("dd", "d"):
            out.append(f"{dt.day:0{len(t)}d}"); continue
        if t in ("mmmmm", "mmmm", "mmm"):
            out.append({"mmmmm": dt.strftime("%b")[0], "mmmm": dt.strftime("%B"),
                        "mmm": dt.strftime("%b")}[t]); continue
        if t in ("m", "mm"):
            # 直前が時、直後が秒なら「分」
            prev = next((low[k] for k in range(idx - 1, -1, -1)
                         if not low[k].startswith('"')), "")
            nxt = next((low[k] for k in range(idx + 1, len(low))
                        if not low[k].startswith('"')), "")
            is_minute = prev in ("h", "hh", "[h]") or nxt in ("s", "ss")
            v = dt.minute if is_minute else dt.month
            out.append(f"{v:0{len(t)}d}"); continue
        if t in ("h", "hh"):
            hr = dt.hour
            if ampm:
                hr = hr % 12 or 12
            out.append(f"{hr:0{len(t)}d}"); continue
        if t in ("s", "ss"):
            out.append(f"{dt.second:0{len(t)}d}"); continue
        if t.startswith(".0"):
            out.append("." + f"{dt.microsecond:06d}"[:len(t) - 1]); continue
        if t == "am/pm":
            out.append("AM" if dt.hour < 12 else "PM"); continue
        if t == "a/p":
            out.append("A" if dt.hour < 12 else "P"); continue
        out.append(tok)
    return "".join(out)


def _strip_brackets_keep_elapsed(fmt: str) -> str:
    return re.sub(r"\[(?![hms]+\])[^\]]*\]", "", fmt, flags=re.IGNORECASE)


def _era(d: date):
    for start, full, short, alpha in ERAS:
        if d >= start:
            return (start, full, short, alpha)
    return (date(1, 1, 1), "西暦", "西", "A")


def _era_year(d: date) -> int:
    start = _era(d)[0]
    y = d.year - start.year + 1
    return y


def format_cell_value(value, number_format: str) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time, timedelta)):
        return format_datetime(value, number_format)
    if isinstance(value, (int, float)):
        return format_number(float(value), number_format)
    return str(value)


# ---------------------------------------------------------------------------
# 罫線
# ---------------------------------------------------------------------------

BORDER_STYLE_MAP = {
    "thin": (1, "solid"),
    "hair": (1, "solid"),
    "medium": (2, "solid"),
    "thick": (3, "solid"),
    "double": (3, "double"),
    "dotted": (1, "dotted"),
    "dashed": (1, "dashed"),
    "dashDot": (1, "dashed"),
    "dashDotDot": (1, "dotted"),
    "mediumDashed": (2, "dashed"),
    "mediumDashDot": (2, "dashed"),
    "mediumDashDotDot": (2, "dotted"),
    "slantDashDot": (1, "dashed"),
}


def border_css(side, resolver: ColorResolver) -> str | None:
    if side is None or not side.style:
        return None
    w, st = BORDER_STYLE_MAP.get(side.style, (1, "solid"))
    color = resolver.resolve(side.color) or "#000000"
    if side.style == "hair":
        color = mix(color, "#ffffff", 0.55)
    return f"{w}px {st} {color}"


# ---------------------------------------------------------------------------
# 描画オブジェクト（画像・図形）
# ---------------------------------------------------------------------------

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
        ".bmp": "image/bmp", ".svg": "image/svg+xml", ".webp": "image/webp",
        ".emf": "image/emf", ".wmf": "image/wmf"}


def _rels_path(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _read_rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    p = _rels_path(part)
    try:
        root = ET.fromstring(z.read(p))
    except (KeyError, ET.ParseError):
        return {}
    base = posixpath.dirname(part)
    out = {}
    for rel in root.findall(f"{{{PKG_REL}}}Relationship"):
        tgt = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            out[rel.get("Id")] = tgt
            continue
        if tgt.startswith("/"):
            out[rel.get("Id")] = tgt.lstrip("/")
        else:
            out[rel.get("Id")] = posixpath.normpath(posixpath.join(base, tgt))
    return out


class DrawingParser:
    """xl/drawings/*.xml を解析して、シートごとの図形リストを返す。"""

    def __init__(self, path: str, resolver: ColorResolver):
        self.resolver = resolver
        self.zip = zipfile.ZipFile(path)
        self.sheet_parts = self._sheet_parts()

    def close(self):
        self.zip.close()

    def _sheet_parts(self) -> dict[str, str]:
        """シート名 → worksheet パート名"""
        try:
            root = ET.fromstring(self.zip.read("xl/workbook.xml"))
        except (KeyError, ET.ParseError):
            return {}
        rels = _read_rels(self.zip, "xl/workbook.xml")
        out = {}
        for sh in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
            rid = sh.get(f"{{{REL}}}id")
            if rid and rid in rels:
                out[sh.get("name")] = rels[rid]
        return out

    def shapes_for(self, sheet_name: str, grid) -> list[dict]:
        part = self.sheet_parts.get(sheet_name)
        if not part:
            return []
        try:
            root = ET.fromstring(self.zip.read(part))
        except (KeyError, ET.ParseError):
            return []
        drels = _read_rels(self.zip, part)
        shapes = []
        for tag in (f"{{{MAIN}}}drawing", f"{{{MAIN}}}legacyDrawing"):
            for d in root.findall(tag):
                rid = d.get(f"{{{REL}}}id")
                target = drels.get(rid)
                if not target or "drawings/drawing" not in target:
                    continue
                shapes.extend(self._parse_drawing(target, grid))
        return shapes

    # -- 個々の drawing パート -------------------------------------------------
    def _parse_drawing(self, part: str, grid) -> list[dict]:
        try:
            root = ET.fromstring(self.zip.read(part))
        except (KeyError, ET.ParseError):
            return []
        rels = _read_rels(self.zip, part)
        out = []
        for anchor in root:
            tag = anchor.tag.split("}")[-1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            box = self._anchor_box(anchor, tag, grid)
            if box is None:
                continue
            for child in anchor:
                ctag = child.tag.split("}")[-1]
                if ctag in ("from", "to", "ext", "pos", "clientData"):
                    continue
                out.extend(self._emit(child, box, rels, part))
        return out

    def _anchor_box(self, anchor, tag, grid):
        def marker(el):
            if el is None:
                return None
            def g(n, default=0):
                e = el.find(f"{{{XDR}}}{n}")
                return int(e.text) if e is not None and e.text else default
            col, coff, row, roff = g("col"), g("colOff"), g("row"), g("rowOff")
            return grid.x_at(col) + emu_to_px(coff), grid.y_at(row) + emu_to_px(roff)

        if tag == "absoluteAnchor":
            pos = anchor.find(f"{{{XDR}}}pos")
            ext = anchor.find(f"{{{XDR}}}ext")
            if pos is None or ext is None:
                return None
            x, y = emu_to_px(pos.get("x")), emu_to_px(pos.get("y"))
            return x, y, emu_to_px(ext.get("cx")), emu_to_px(ext.get("cy"))
        frm = marker(anchor.find(f"{{{XDR}}}from"))
        if frm is None:
            return None
        if tag == "twoCellAnchor":
            to = marker(anchor.find(f"{{{XDR}}}to"))
            if to is None:
                return None
            return frm[0], frm[1], max(0.0, to[0] - frm[0]), max(0.0, to[1] - frm[1])
        ext = anchor.find(f"{{{XDR}}}ext")
        cx = emu_to_px(ext.get("cx")) if ext is not None else 0
        cy = emu_to_px(ext.get("cy")) if ext is not None else 0
        return frm[0], frm[1], cx, cy

    def _emit(self, node, box, rels, part) -> list[dict]:
        tag = node.tag.split("}")[-1]
        if tag == "pic":
            s = self._picture(node, box, rels)
            return [s] if s else []
        if tag in ("sp", "cxnSp"):
            return [self._shape(node, box)]
        if tag == "grpSp":
            return self._group(node, box, rels, part)
        if tag == "graphicFrame":      # グラフ・SmartArt 等（枠のみ）
            return []
        return []

    def _group(self, node, box, rels, part) -> list[dict]:
        xfrm = node.find(f"{{{XDR}}}grpSpPr/{{{A_NS}}}xfrm")
        out = []
        if xfrm is None:
            for ch in node:
                if ch.tag.split("}")[-1] in ("sp", "pic", "grpSp", "cxnSp"):
                    out.extend(self._emit(ch, box, rels, part))
            return out
        off = xfrm.find(f"{{{A_NS}}}off")
        ext = xfrm.find(f"{{{A_NS}}}ext")
        choff = xfrm.find(f"{{{A_NS}}}chOff")
        chext = xfrm.find(f"{{{A_NS}}}chExt")
        gx, gy, gw, gh = box
        cox = emu_to_px(choff.get("x")) if choff is not None else 0
        coy = emu_to_px(choff.get("y")) if choff is not None else 0
        cw = emu_to_px(chext.get("cx")) if chext is not None else (gw or 1)
        chh = emu_to_px(chext.get("cy")) if chext is not None else (gh or 1)
        sx = (gw / cw) if cw else 1.0
        sy = (gh / chh) if chh else 1.0
        for ch in node:
            ctag = ch.tag.split("}")[-1]
            if ctag not in ("sp", "pic", "grpSp", "cxnSp"):
                continue
            cxf = ch.find(f".//{{{A_NS}}}xfrm")
            if cxf is None:
                out.extend(self._emit(ch, box, rels, part))
                continue
            coff = cxf.find(f"{{{A_NS}}}off")
            cext = cxf.find(f"{{{A_NS}}}ext")
            x = gx + (emu_to_px(coff.get("x")) - cox) * sx if coff is not None else gx
            y = gy + (emu_to_px(coff.get("y")) - coy) * sy if coff is not None else gy
            w = emu_to_px(cext.get("cx")) * sx if cext is not None else gw
            h = emu_to_px(cext.get("cy")) * sy if cext is not None else gh
            out.extend(self._emit(ch, (x, y, w, h), rels, part))
        return out

    def _picture(self, node, box, rels):
        blip = node.find(f".//{{{A_NS}}}blip")
        if blip is None:
            return None
        rid = blip.get(f"{{{REL}}}embed") or blip.get(f"{{{REL}}}link")
        target = rels.get(rid)
        if not target:
            return None
        try:
            data = self.zip.read(target)
        except KeyError:
            return None
        ext = os.path.splitext(target)[1].lower()
        mime = MIME.get(ext, "application/octet-stream")
        if ext in (".emf", ".wmf"):
            return None            # ブラウザで描けない形式は省略
        uri = f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
        rot = self._rotation(node)
        name = ""
        nv = node.find(f"{{{XDR}}}nvPicPr/{{{XDR}}}cNvPr")
        if nv is not None:
            name = nv.get("descr") or nv.get("name") or ""
        return {"kind": "img", "box": box, "src": uri, "rot": rot, "alt": name}

    def _rotation(self, node) -> float:
        xfrm = node.find(f".//{{{A_NS}}}xfrm")
        if xfrm is None:
            return 0.0
        rot = xfrm.get("rot")
        return (int(rot) / 60000.0) if rot else 0.0

    def _flips(self, node):
        xfrm = node.find(f".//{{{A_NS}}}xfrm")
        if xfrm is None:
            return False, False
        return xfrm.get("flipH") in ("1", "true"), xfrm.get("flipV") in ("1", "true")

    def _color_of(self, el) -> str | None:
        """a:solidFill 等の色要素 → #rrggbb"""
        if el is None:
            return None
        srgb = el.find(f".//{{{A_NS}}}srgbClr")
        if srgb is not None and srgb.get("val"):
            c = "#" + srgb.get("val").lower()
            return self._modify(c, srgb)
        scheme = el.find(f".//{{{A_NS}}}schemeClr")
        if scheme is not None:
            name = scheme.get("val")
            alias = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
                     "phClr": "accent1"}.get(name, name)
            rgb = self.resolver.theme.get(alias)
            if rgb:
                return self._modify("#" + rgb.lower(), scheme)
        sysc = el.find(f".//{{{A_NS}}}sysClr")
        if sysc is not None:
            return "#" + (sysc.get("lastClr") or "000000").lower()
        return None

    def _modify(self, color: str, node) -> str:
        """lumMod / lumOff / shade / tint を近似適用。"""
        def val(tag):
            e = node.find(f"{{{A_NS}}}{tag}")
            return int(e.get("val")) / 100000.0 if e is not None and e.get("val") else None
        lum_mod, lum_off = val("lumMod"), val("lumOff")
        shade, tint = val("shade"), val("tint")
        r, g, b = (int(color[i:i + 2], 16) / 255.0 for i in (1, 3, 5))
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if lum_mod is not None:
            l *= lum_mod
        if lum_off is not None:
            l += lum_off
        if shade is not None:
            l *= shade
        if tint is not None:
            l = l * tint + (1 - tint)
        l = min(1.0, max(0.0, l))
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return "#%02x%02x%02x" % (round(r * 255), round(g * 255), round(b * 255))

    def _shape(self, node, box) -> dict:
        sp_pr = node.find(f"{{{XDR}}}spPr")
        prst = "rect"
        if sp_pr is not None:
            geom = sp_pr.find(f"{{{A_NS}}}prstGeom")
            if geom is not None:
                prst = geom.get("prst") or "rect"
            elif sp_pr.find(f"{{{A_NS}}}custGeom") is not None:
                prst = "rect"
        # 塗り
        fill = None
        if sp_pr is not None:
            if sp_pr.find(f"{{{A_NS}}}noFill") is not None:
                fill = None
            else:
                sf = sp_pr.find(f"{{{A_NS}}}solidFill")
                if sf is not None:
                    fill = self._color_of(sf)
                else:
                    gf = sp_pr.find(f"{{{A_NS}}}gradFill")
                    if gf is not None:
                        fill = self._color_of(gf)
        has_explicit_fill = sp_pr is not None and (
            sp_pr.find(f"{{{A_NS}}}noFill") is not None
            or sp_pr.find(f"{{{A_NS}}}solidFill") is not None
            or sp_pr.find(f"{{{A_NS}}}gradFill") is not None)
        # 線
        line_color, line_w, line_dash = None, 1.0, None
        ln = sp_pr.find(f"{{{A_NS}}}ln") if sp_pr is not None else None
        has_explicit_line = False
        head_end = tail_end = None
        if ln is not None:
            has_explicit_line = True
            if ln.get("w"):
                line_w = emu_to_px(ln.get("w"))
            if ln.find(f"{{{A_NS}}}noFill") is not None:
                line_color = None
            else:
                sf = ln.find(f"{{{A_NS}}}solidFill")
                line_color = self._color_of(sf) if sf is not None else None
                if sf is None:
                    has_explicit_line = False
            d = ln.find(f"{{{A_NS}}}prstDash")
            if d is not None:
                line_dash = d.get("val")
            he = ln.find(f"{{{A_NS}}}headEnd")
            te = ln.find(f"{{{A_NS}}}tailEnd")
            head_end = he.get("type") if he is not None else None
            tail_end = te.get("type") if te is not None else None
        # スタイル参照（既定色）
        style = node.find(f"{{{XDR}}}style")
        if style is not None:
            if not has_explicit_fill:
                fill = self._color_of(style.find(f"{{{A_NS}}}fillRef"))
            if not has_explicit_line and line_color is None:
                line_color = self._color_of(style.find(f"{{{A_NS}}}lnRef"))
        if not has_explicit_fill and fill is None and style is None:
            fill = None
        flip_h, flip_v = self._flips(node)
        return {
            "kind": "shape", "box": box, "prst": prst, "fill": fill,
            "line": line_color, "line_w": max(line_w, 0.5) if line_color else 0,
            "dash": line_dash, "rot": self._rotation(node),
            "flip_h": flip_h, "flip_v": flip_v,
            "head": head_end, "tail": tail_end,
            "text": self._text_body(node),
        }

    def _text_body(self, node) -> dict | None:
        tx = node.find(f"{{{XDR}}}txBody")
        if tx is None:
            return None
        paras = []
        for p in tx.findall(f"{{{A_NS}}}p"):
            ppr = p.find(f"{{{A_NS}}}pPr")
            algn = ppr.get("algn") if ppr is not None else None
            runs = []
            for r in p:
                t = r.tag.split("}")[-1]
                if t == "br":
                    runs.append({"text": "\n"})
                    continue
                if t not in ("r", "fld"):
                    continue
                tnode = r.find(f"{{{A_NS}}}t")
                if tnode is None or tnode.text is None:
                    continue
                rpr = r.find(f"{{{A_NS}}}rPr")
                run = {"text": tnode.text}
                if rpr is not None:
                    if rpr.get("sz"):
                        run["size"] = int(rpr.get("sz")) / 100.0
                    run["b"] = rpr.get("b") in ("1", "true")
                    run["i"] = rpr.get("i") in ("1", "true")
                    run["u"] = rpr.get("u") not in (None, "none")
                    c = self._color_of(rpr.find(f"{{{A_NS}}}solidFill"))
                    if c:
                        run["color"] = c
                    latin = rpr.find(f"{{{A_NS}}}latin")
                    ea = rpr.find(f"{{{A_NS}}}ea")
                    face = (ea.get("typeface") if ea is not None else None) or \
                           (latin.get("typeface") if latin is not None else None)
                    if face and not face.startswith("+"):
                        run["font"] = face
                runs.append(run)
            if runs:
                paras.append({"algn": algn, "runs": runs})
        if not paras:
            return None
        body = tx.find(f"{{{A_NS}}}bodyPr")
        anchor = body.get("anchor") if body is not None else None
        vert = body.get("vert") if body is not None else None
        return {"paras": paras, "anchor": anchor, "vert": vert}


# 図形 → SVG パス（w,h は 100 正規化ではなく実寸で生成）
def shape_path(prst: str, w: float, h: float) -> str | None:
    W, H = w, h
    a = min(W, H) * 0.35          # 矢印の軸太さ相当
    if prst in ("line", "straightConnector1"):
        return f"M0,0 L{W},{H}"
    if prst in ("triangle", "isoscelesTriangle"):
        return f"M{W/2},0 L{W},{H} L0,{H} Z"
    if prst == "rtTriangle":
        return f"M0,0 L0,{H} L{W},{H} Z"
    if prst == "diamond":
        return f"M{W/2},0 L{W},{H/2} L{W/2},{H} L0,{H/2} Z"
    if prst == "parallelogram":
        d = W * 0.25
        return f"M{d},0 L{W},0 L{W-d},{H} L0,{H} Z"
    if prst == "trapezoid":
        d = W * 0.25
        return f"M{d},0 L{W-d},0 L{W},{H} L0,{H} Z"
    if prst == "hexagon":
        d = W * 0.25
        return f"M{d},0 L{W-d},0 L{W},{H/2} L{W-d},{H} L{d},{H} L0,{H/2} Z"
    if prst == "pentagon" or prst == "homePlate":
        d = W * 0.25
        return f"M0,0 L{W-d},0 L{W},{H/2} L{W-d},{H} L0,{H} Z"
    if prst == "chevron":
        d = W * 0.25
        return f"M0,0 L{W-d},0 L{W},{H/2} L{W-d},{H} L0,{H} L{d},{H/2} Z"
    if prst == "rightArrow":
        return (f"M0,{H/2-a/2} L{W-H*0.4},{H/2-a/2} L{W-H*0.4},0 L{W},{H/2} "
                f"L{W-H*0.4},{H} L{W-H*0.4},{H/2+a/2} L0,{H/2+a/2} Z")
    if prst == "leftArrow":
        return (f"M{W},{H/2-a/2} L{H*0.4},{H/2-a/2} L{H*0.4},0 L0,{H/2} "
                f"L{H*0.4},{H} L{H*0.4},{H/2+a/2} L{W},{H/2+a/2} Z")
    if prst == "downArrow":
        b = min(W, H) * 0.35
        return (f"M{W/2-b/2},0 L{W/2+b/2},0 L{W/2+b/2},{H-W*0.4} L{W},{H-W*0.4} "
                f"L{W/2},{H} L0,{H-W*0.4} L{W/2-b/2},{H-W*0.4} Z")
    if prst == "upArrow":
        b = min(W, H) * 0.35
        return (f"M{W/2},0 L{W},{W*0.4} L{W/2+b/2},{W*0.4} L{W/2+b/2},{H} "
                f"L{W/2-b/2},{H} L{W/2-b/2},{W*0.4} L0,{W*0.4} Z")
    if prst == "leftRightArrow":
        return (f"M0,{H/2} L{H*0.4},0 L{H*0.4},{H/2-a/2} L{W-H*0.4},{H/2-a/2} "
                f"L{W-H*0.4},0 L{W},{H/2} L{W-H*0.4},{H} L{W-H*0.4},{H/2+a/2} "
                f"L{H*0.4},{H/2+a/2} L{H*0.4},{H} Z")
    if prst == "star5":
        pts = []
        for i in range(10):
            ang = -math.pi / 2 + i * math.pi / 5
            r = 0.5 if i % 2 == 0 else 0.2
            pts.append(f"{W/2 + math.cos(ang)*W*r},{H/2 + math.sin(ang)*H*r}")
        return "M" + " L".join(pts) + " Z"
    if prst.startswith("bentConnector") or prst.startswith("curvedConnector"):
        return f"M0,0 L{W/2},0 L{W/2},{H} L{W},{H}"
    return None


# ---------------------------------------------------------------------------
# グリッド（列幅・行高）
# ---------------------------------------------------------------------------

class Grid:
    def __init__(self, ws, max_col: int, max_row: int):
        self.ws = ws
        self.max_col = max_col
        self.max_row = max_row
        fmt = ws.sheet_format
        self.default_col = fmt.defaultColWidth or DEFAULT_COL_WIDTH
        self.default_row_pt = fmt.defaultRowHeight or DEFAULT_ROW_HEIGHT_PT

        widths: dict[int, float] = {}
        hidden_cols: set[int] = set()
        for key, cd in ws.column_dimensions.items():
            lo = getattr(cd, "min", None)
            hi = getattr(cd, "max", None)
            if not lo:
                continue
            hi = hi or lo
            for c in range(lo, min(hi, max_col) + 1):
                if cd.hidden:
                    hidden_cols.add(c)
                    widths[c] = 0.0
                elif cd.width is not None:
                    widths[c] = cd.width
        self.hidden_cols = hidden_cols
        self.col_px = [0] * (max_col + 1)
        for c in range(1, max_col + 1):
            self.col_px[c] = 0 if c in hidden_cols else col_width_to_px(widths.get(c, self.default_col))

        self.hidden_rows: set[int] = set()
        self.row_px = [0.0] * (max_row + 1)
        for r in range(1, max_row + 1):
            rd = ws.row_dimensions.get(r)
            if rd is not None and rd.hidden:
                self.hidden_rows.add(r)
                self.row_px[r] = 0.0
            elif rd is not None and rd.height is not None:
                self.row_px[r] = pt_to_px(rd.height)
            else:
                self.row_px[r] = pt_to_px(self.default_row_pt)

        self._x = [0.0] * (max_col + 2)
        for c in range(1, max_col + 2):
            self._x[c] = self._x[c - 1] + (self.col_px[c - 1] if c - 1 >= 1 else 0)
        self._y = [0.0] * (max_row + 2)
        for r in range(1, max_row + 2):
            self._y[r] = self._y[r - 1] + (self.row_px[r - 1] if r - 1 >= 1 else 0)

    @property
    def width(self) -> float:
        return sum(self.col_px[1:])

    @property
    def height(self) -> float:
        return sum(self.row_px[1:])

    def x_at(self, col0: int) -> float:
        """0 始まり列インデックスの左端 x。使用範囲の外は列定義を見て外挿する。"""
        c = col0 + 1
        if c <= self.max_col + 1:
            return self._x[c]
        x = self._x[self.max_col + 1]
        for cc in range(self.max_col + 1, c):
            cd = self.ws.column_dimensions.get(get_column_letter(cc))
            w = cd.width if (cd is not None and cd.width) else self.default_col
            x += 0 if (cd is not None and cd.hidden) else col_width_to_px(w)
        return x

    def y_at(self, row0: int) -> float:
        r = row0 + 1
        if r <= self.max_row + 1:
            return self._y[r]
        y = self._y[self.max_row + 1]
        for rr in range(self.max_row + 1, r):
            rd = self.ws.row_dimensions.get(rr)
            h = rd.height if (rd is not None and rd.height) else self.default_row_pt
            y += 0 if (rd is not None and rd.hidden) else pt_to_px(h)
        return y


# ---------------------------------------------------------------------------
# セルのスタイル → CSS
# ---------------------------------------------------------------------------

FONT_FALLBACK = ('"Meiryo","Yu Gothic","MS PGothic","Hiragino Kaku Gothic ProN",'
                 '"Noto Sans JP",sans-serif')

H_ALIGN = {"left": "left", "center": "center", "right": "right",
           "justify": "justify", "distributed": "justify",
           "centerContinuous": "center", "fill": "left"}
V_ALIGN = {"top": "top", "center": "middle", "bottom": "bottom",
           "justify": "middle", "distributed": "middle"}


class StyleSheet:
    """同一書式のセルで CSS クラスを共有し、出力サイズを抑える。"""

    def __init__(self, resolver: ColorResolver, gridlines: bool):
        self.resolver = resolver
        self.gridlines = gridlines
        self._by_decl: dict[tuple[str, str], str] = {}
        self._by_style_key: dict[tuple, str] = {}
        self.rules: list[tuple[str, str, str]] = []

    def class_for(self, cell, grid_r: bool = True, grid_b: bool = True) -> str:
        try:
            key = (tuple(cell._style), grid_r, grid_b)
        except Exception:
            key = None
        if key is not None and key in self._by_style_key:
            return self._by_style_key[key]
        td, sp = self._build(cell, grid_r, grid_b)
        name = self._by_decl.get((td, sp))
        if name is None:
            name = f"s{len(self._by_decl)}"
            self._by_decl[(td, sp)] = name
            self.rules.append((name, td, sp))
        if key is not None:
            self._by_style_key[key] = name
        return name

    def _build(self, cell, grid_r: bool = True, grid_b: bool = True) -> tuple[str, str]:
        R = self.resolver
        font, fill, border, al = cell.font, cell.fill, cell.border, cell.alignment
        td: list[str] = []
        sp: list[str] = []

        # --- 塗り ---
        bg = None
        if fill is not None and fill.fill_type:
            fg = R.resolve(fill.fgColor)
            bgc = R.resolve(fill.bgColor) or "#ffffff"
            if fill.fill_type == "solid":
                bg = fg
            elif fill.fill_type.startswith("gray") or fill.fill_type in (
                    "darkGray", "mediumGray", "lightGray", "gray125", "gray0625"):
                ratio = {"darkGray": .75, "mediumGray": .5, "lightGray": .25,
                         "gray125": .125, "gray0625": .0625}.get(fill.fill_type, .25)
                if fg:
                    bg = mix(fg, bgc, ratio)
            elif fg:
                bg = mix(fg, bgc, 0.5)
        if bg:
            td.append(f"background:{bg}")

        # --- 罫線 ---
        for side, prop in (("top", "border-top"), ("bottom", "border-bottom"),
                           ("left", "border-left"), ("right", "border-right")):
            css = border_css(getattr(border, side, None), R) if border else None
            if css:
                td.append(f"{prop}:{css}")
            elif self.gridlines and ((prop == "border-right" and grid_r)
                                     or (prop == "border-bottom" and grid_b)):
                # 色を CSS 変数にしておき、画面の切り替えと印刷時の非表示を効かせる。
                # 幅は 1px のまま残るので、消しても方眼の位置はずれない。
                # 隣のセルが罫線を持つ辺には出さない（grid_r / grid_b）。border-collapse は
                # 太さが同じなら左・上側の線を採用するため、目盛線が隣の罫線を消してしまう。
                td.append(f"{prop}:1px solid var(--gl)")
        # 斜線
        if border is not None:
            diag_css = []
            dcolor = R.resolve(border.diagonal.color) if border.diagonal else None
            dcolor = dcolor or "#000000"
            if border.diagonal is not None and border.diagonal.style:
                if border.diagonalDown:
                    diag_css.append(f"linear-gradient(to bottom right,transparent calc(50% - 0.5px),"
                                    f"{dcolor} calc(50% - 0.5px),{dcolor} calc(50% + 0.5px),"
                                    f"transparent calc(50% + 0.5px))")
                if border.diagonalUp:
                    diag_css.append(f"linear-gradient(to top right,transparent calc(50% - 0.5px),"
                                    f"{dcolor} calc(50% - 0.5px),{dcolor} calc(50% + 0.5px),"
                                    f"transparent calc(50% + 0.5px))")
            if diag_css:
                td.append("background-image:" + ",".join(diag_css))

        # --- フォント ---
        if font is not None:
            if font.name:
                td.append(f'font-family:"{font.name}",{FONT_FALLBACK}')
            if font.sz:
                td.append(f"font-size:{font.sz}pt")
            if font.b:
                td.append("font-weight:700")
            if font.i:
                td.append("font-style:italic")
            deco = []
            if font.u:
                deco.append("underline")
            if font.strike:
                deco.append("line-through")
            if deco:
                td.append("text-decoration:" + " ".join(deco))
                if font.u in ("double", "doubleAccounting"):
                    td.append("text-decoration-style:double")
            c = R.resolve(font.color)
            if c:
                td.append(f"color:{c}")
            if font.vertAlign == "superscript":
                sp.append("vertical-align:super;font-size:0.7em")
            elif font.vertAlign == "subscript":
                sp.append("vertical-align:sub;font-size:0.7em")

        # --- 配置 ---
        h = H_ALIGN.get(al.horizontal or "", None) if al else None
        v = V_ALIGN.get(al.vertical or "", "bottom") if al else "bottom"
        wrap = bool(al and al.wrapText)
        rot = (al.textRotation if al else 0) or 0
        indent = (al.indent if al else 0) or 0
        transforms: list[str] = []

        td.append(f"vertical-align:{v}")
        if h:
            td.append(f"text-align:{h}")
        if wrap:
            sp.append("white-space:pre-wrap")
            sp.append("position:static;display:block;width:100%")
            if h:
                sp.append(f"text-align:{h}")
            if indent:
                sp.append(f"padding-left:{indent * INDENT_PX}px")
        else:
            sp.append("white-space:pre")
            sp.append("position:absolute")
            # 横方向のアンカー（左右の余白 CELL_PAD は td>i の padding が担う）
            if h == "center":
                sp.append("left:50%")
                transforms.append("translateX(-50%)")
            elif h == "right":
                sp.append(f"right:{indent * INDENT_PX}px")
            else:
                sp.append(f"left:{indent * INDENT_PX}px")
            # 縦方向のアンカー
            if v == "top":
                sp.append("top:0")
            elif v == "middle":
                sp.append("top:50%")
                transforms.append("translateY(-50%)")
            else:
                sp.append("bottom:0")

        if rot:
            if rot == 255:
                sp.append("writing-mode:vertical-rl;text-orientation:upright;"
                          "white-space:pre;letter-spacing:0.05em")
            else:
                deg = -rot if rot <= 90 else (rot - 90)
                transforms.append(f"rotate({deg}deg)")
                sp.append("transform-origin:center center")
        if transforms:
            sp.append("transform:" + " ".join(transforms))
        return ";".join(td), ";".join(sp)

    def css(self) -> str:
        out = []
        for name, td, sp in self.rules:
            if td:
                out.append(f".{name}{{{td}}}")
            if sp:
                out.append(f".{name}>i{{{sp}}}")
        return "\n".join(out)


# ---------------------------------------------------------------------------
# シート → HTML
# ---------------------------------------------------------------------------

def rich_text_html(value) -> str | None:
    """CellRichText（セル内の部分書式）を span 列に変換。"""
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
    except ImportError:
        return None
    if not isinstance(value, CellRichText):
        return None
    parts = []
    for item in value:
        if isinstance(item, str):
            parts.append(esc(item))
            continue
        if isinstance(item, TextBlock):
            f = item.font
            css = []
            if f is not None:
                if f.b:
                    css.append("font-weight:700")
                if f.i:
                    css.append("font-style:italic")
                if f.u:
                    css.append("text-decoration:underline")
                if f.sz:
                    css.append(f"font-size:{f.sz}pt")
                if f.rFont:
                    css.append(f'font-family:"{f.rFont}",{FONT_FALLBACK}')
                col = getattr(f, "color", None)
                if col is not None and getattr(col, "rgb", None) and isinstance(col.rgb, str):
                    css.append("color:#" + col.rgb[-6:].lower())
            parts.append(f'<span style="{";".join(css)}">{esc(str(item.text))}</span>'
                         if css else esc(str(item.text)))
    return "".join(parts)


def load_formulas(src: str, names: list[str]) -> dict[str, dict[tuple[int, int], str]]:
    """数式だけを高速に読み出す（read_only なので書式は読まない）。"""
    try:
        from openpyxl.worksheet.formula import ArrayFormula
    except ImportError:
        ArrayFormula = ()
    out: dict[str, dict[tuple[int, int], str]] = {}
    wb = load_workbook(src, data_only=False, read_only=True)
    try:
        for name in names:
            if name not in wb.sheetnames:
                continue
            m: dict[tuple[int, int], str] = {}
            for row in wb[name].iter_rows():
                for c in row:
                    v = c.value
                    if ArrayFormula and isinstance(v, ArrayFormula):
                        v = v.text
                    if isinstance(v, str) and v.startswith("="):
                        m[(c.row, c.column)] = v
            out[name] = m
    finally:
        wb.close()
    return out


def tidy_formula(f: str) -> str:
    """_xlfn.XLOOKUP のような内部表記を Excel の表示名に戻す。"""
    return re.sub(r"_xlfn\.(_xlws\.)?", "", f)


def render_sheet(ws, resolver: ColorResolver, drawings: list[dict],
                 gridlines: bool, sheet_id: str, opts,
                 formulas: dict[tuple[int, int], str] | None = None,
                 bounds: tuple[int, int, int, int] | None = None,
                 page_breaks: set[int] | None = None,
                 title_rows: tuple[int, int] | None = None) -> dict:
    full_row = min(ws.max_row or 1, opts.max_rows)
    full_col = min(ws.max_column or 1, opts.max_cols)
    grid = Grid(ws, full_col, full_row)
    min_row, min_col, max_row, max_col = bounds or (1, 1, full_row, full_col)
    max_row, max_col = min(max_row, full_row), min(max_col, full_col)
    styles = StyleSheet(resolver, gridlines)
    page_breaks = page_breaks or set()

    # 結合セル（印刷範囲で切った場合は、はみ出した分を詰めて表示する）
    covered: dict[tuple[int, int], tuple[int, int]] = {}
    spans: dict[tuple[int, int], tuple[int, int]] = {}
    src: dict[tuple[int, int], tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if r1 > max_row or c1 > max_col or r2 < min_row or c2 < min_col:
            continue
        ar, ac = max(r1, min_row), max(c1, min_col)
        er, ec = min(r2, max_row), min(c2, max_col)
        spans[(ar, ac)] = (er - ar + 1, ec - ac + 1)
        if (ar, ac) != (r1, c1):
            src[(ar, ac)] = (r1, c1)
        for r in range(ar, er + 1):
            for c in range(ac, ec + 1):
                if (r, c) != (ar, ac):
                    covered[(r, c)] = (ar, ac)

    # 1パス目: 表示文字列を確定し、行ごとの「文字が入っているセル」を把握する
    #（Excel は隣が空セルのときだけ文字をはみ出させるため、その判定に使う）
    content: dict[tuple[int, int], str] = {}
    occupied: dict[int, set[int]] = {}
    formulas = formulas or {}
    missing_cache = 0
    for r in range(min_row, max_row + 1):
        occ = set()
        for c in range(min_col, max_col + 1):
            if (r, c) in covered:
                occ.add(c)
                continue
            sr, sc = src.get((r, c), (r, c))
            cell = ws.cell(row=sr, column=sc)
            if cell.value is None:
                # 数式なのに計算結果が保存されていないセル
                fml = formulas.get((r, c))
                if not fml:
                    continue
                missing_cache += 1
                if not opts.show_formula:
                    continue
                content[(r, c)] = esc(tidy_formula(fml))
                occ.add(c)
                rs, cs = spans.get((r, c), (1, 1))
                for cc in range(c, min(c + cs, max_col + 1)):
                    occ.add(cc)
                continue
            hval = rich_text_html(cell.value)
            if hval is None:
                text = format_cell_value(cell.value, cell.number_format)
                hval = esc(text) if text != "" else ""
            if hval == "":
                continue
            content[(r, c)] = hval
            occ.add(c)
            rs, cs = spans.get((r, c), (1, 1))
            for cc in range(c, min(c + cs, max_col + 1)):
                occ.add(cc)
        occupied[r] = occ

    def has_border(r: int, c: int, side: str) -> bool:
        if not (min_row <= r <= max_row and min_col <= c <= max_col):
            return False
        b = ws.cell(row=r, column=c).border
        s = getattr(b, side, None) if b is not None else None
        return bool(s is not None and s.style)

    def grid_edges(r: int, c: int, rs: int, cs: int) -> tuple[bool, bool]:
        """右辺・下辺に目盛線を引いてよいか。隣に罫線があるなら引かない。"""
        if not gridlines:
            return True, True
        rb = r + rs
        while rb in grid.hidden_rows:            # 非表示行は出力しないので、その先を見る
            rb += 1
        return (not any(has_border(rr, c + cs, "left") for rr in range(r, r + rs)),
                not any(has_border(rb, cc, "top") for cc in range(c, c + cs)))

    def free_width(r: int, c_from: int, step: int) -> tuple[float, bool]:
        """隣接する空セルの合計幅と、行端まで空だったかを返す。"""
        w = 0.0
        occ = occupied.get(r, ())
        c = c_from
        while min_col <= c <= max_col:
            if c in occ:
                return w, True
            w += grid.col_px[c]
            c += step
        return w, False

    # 印刷範囲で切った場合は、その左上を原点にする
    x0, y0 = grid.x_at(min_col - 1), grid.y_at(min_row - 1)
    table_w = sum(grid.col_px[min_col:max_col + 1])
    table_h = sum(grid.row_px[min_row:max_row + 1])

    out = []
    cw, ch = table_w, table_h
    shapes = []
    for s in drawings:                     # 図形が用紙外に出る場合は台紙を広げる
        x, y, w, h = s["box"]
        x, y = x - x0, y - y0
        if x + w < 0 or y + h < 0:
            continue
        s = dict(s, box=(x, y, w, h))
        shapes.append(s)
        cw = max(cw, x + w + 2)
        ch = max(ch, y + h + 2)
    out.append(f'<div class="canvas" style="width:{cw:.0f}px;height:{ch:.0f}px">')
    out.append(f'<table style="width:{table_w:.0f}px"><colgroup>')
    for c in range(min_col, max_col + 1):
        out.append(f'<col style="width:{grid.col_px[c]}px">')
    out.append("</colgroup>")

    thead_end = 0
    if title_rows and title_rows[0] <= min_row:
        thead_end = min(title_rows[1], max_row)
    out.append("<thead>" if thead_end >= min_row else "<tbody>")

    for r in range(min_row, max_row + 1):
        if r in grid.hidden_rows:
            continue
        if thead_end and r == thead_end + 1:
            out.append("</thead><tbody>")
        brk = ' class="pb"' if r in page_breaks and r > min_row else ""
        out.append(f'<tr{brk} style="height:{grid.row_px[r]:.0f}px">')
        for c in range(min_col, max_col + 1):
            if (r, c) in covered:
                continue
            if grid.col_px[c] == 0 and (r, c) not in spans:
                out.append('<td class="hidden"></td>')
                continue
            sr, sc = src.get((r, c), (r, c))
            cell = ws.cell(row=sr, column=sc)
            rs, cs = spans.get((r, c), (1, 1))
            cls = styles.class_for(cell, *grid_edges(r, c, rs, cs))
            attr = f' class="{cls}"'
            if rs > 1:
                attr += f' rowspan="{rs}"'
            if cs > 1:
                attr += f' colspan="{cs}"'
            body = content.get((r, c))
            if not body:
                out.append(f"<td{attr}></td>")
                continue

            # 見切れの再現。Excel と同じく、横は「隣に文字があるところ」まで、
            # 縦は行の高さまでで切る。切った文字も DOM には残すのでコピー・検索できる。
            al = cell.alignment
            wrap = bool(al and al.wrapText)
            halign = (al.horizontal if al else None) or ""
            own = sum(grid.col_px[c:c + cs])
            limit, kls = "", ""
            if wrap:
                own_h = sum(grid.row_px[r:r + rs])
                limit = f' style="max-height:{own_h:.0f}px"'
                kls = " cw"
            elif not (al and al.textRotation):
                if halign in ("", "left", "general", "justify", "distributed", "fill"):
                    avail, blocked = free_width(r, c + cs, 1)
                    if blocked:
                        limit = f' style="max-width:{own + avail:.0f}px"'
                        kls = " c"
                elif halign in ("center", "centerContinuous"):
                    la, lb = free_width(r, c - 1, -1)
                    ra, rb = free_width(r, c + cs, 1)
                    if lb or rb:
                        limit = f' style="max-width:{own + 2 * min(la, ra):.0f}px"'
                        kls = " c"
                elif halign == "right":
                    avail, blocked = free_width(r, c - 1, -1)
                    if blocked:
                        limit = f' style="max-width:{own + avail:.0f}px"'
                        kls = " cr"
            if kls:
                kls = f' class="{kls.strip()}"'

            inner = body
            link = cell.hyperlink
            if link is not None and getattr(link, "target", None):
                inner = f'<a href="{esc(link.target)}">{body}</a>'
            tips = []
            if cell.comment is not None and cell.comment.text:
                tips.append(cell.comment.text)
            if opts.formula_tips:
                fml = formulas.get((sr, sc))
                if fml:
                    tips.append(tidy_formula(fml))
            tip = f' title="{esc(chr(10).join(tips))}"' if tips else ""
            out.append(f"<td{attr}{tip}><i{kls}{limit}>{inner}</i></td>")
        out.append("</tr>")
    out.append("</tbody></table>")

    if shapes:
        out.append('<div class="draw">')
        for s in shapes:
            out.append(render_shape(s))
        out.append("</div>")
    out.append("</div>")
    return {"html": "".join(out), "css": styles.css(), "missing": missing_cache,
            "width": cw, "height": ch, "rows": max_row - min_row + 1,
            "cols": max_col - min_col + 1}


def render_shape(s: dict) -> str:
    x, y, w, h = s["box"]
    w = max(w, 1.0)
    h = max(h, 1.0)
    base = f"left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{h:.1f}px"
    rot = s.get("rot") or 0
    if rot:
        base += f";transform:rotate({rot:.2f}deg)"
    if s["kind"] == "img":
        alt = esc(s.get("alt") or "")
        return (f'<img class="dobj" style="{base}" src="{s["src"]}" alt="{alt}">')

    prst = s.get("prst", "rect")
    path = shape_path(prst, w, h)
    fill, line, lw = s.get("fill"), s.get("line"), s.get("line_w") or 0
    dash = s.get("dash")
    body = []

    is_line = prst in ("line", "straightConnector1") or prst.startswith(("bentConnector",
                                                                        "curvedConnector"))
    if path is not None:
        stroke = line or ("#000000" if is_line else "none")
        sw = lw if lw else (1 if is_line else 0)
        dasharray = ""
        if dash in ("dash", "sysDash", "lgDash"):
            dasharray = f' stroke-dasharray="{max(sw*3,4)},{max(sw*2,3)}"'
        elif dash in ("dot", "sysDot"):
            dasharray = f' stroke-dasharray="{max(sw,1)},{max(sw*2,2)}"'
        elif dash in ("dashDot", "sysDashDot", "lgDashDot"):
            dasharray = f' stroke-dasharray="{max(sw*3,4)},{max(sw*2,3)},{max(sw,1)},{max(sw*2,3)}"'
        # 反転
        tf = ""
        if s.get("flip_h") or s.get("flip_v"):
            sx, sy = (-1 if s.get("flip_h") else 1), (-1 if s.get("flip_v") else 1)
            cx, cy = w / 2, h / 2
            tf = f' transform="translate({cx if sx<0 else 0},{cy if sy<0 else 0}) scale({sx},{sy}) translate({-cx if sx<0 else 0},{-cy if sy<0 else 0})"'
        marker = ""
        defs = ""
        uid = f"m{abs(hash((x, y, w, h, prst))) % 100000}"
        if s.get("tail") and s["tail"] != "none":
            defs += (f'<marker id="{uid}t" markerWidth="6" markerHeight="6" refX="5" refY="3" '
                     f'orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="{line or "#000"}"/></marker>')
            marker += f' marker-end="url(#{uid}t)"'
        if s.get("head") and s["head"] != "none":
            defs += (f'<marker id="{uid}h" markerWidth="6" markerHeight="6" refX="1" refY="3" '
                     f'orient="auto"><path d="M6,0 L0,3 L6,6 Z" fill="{line or "#000"}"/></marker>')
            marker += f' marker-start="url(#{uid}h)"'
        body.append(
            f'<svg class="dsvg" width="{w:.1f}" height="{h:.1f}" viewBox="0 0 {w:.1f} {h:.1f}" '
            f'overflow="visible">' +
            (f"<defs>{defs}</defs>" if defs else "") +
            f'<path d="{path}" fill="{fill or "none"}" stroke="{stroke}" '
            f'stroke-width="{sw}"{dasharray}{marker}{tf}/></svg>')
        shell_style = base
    else:
        # 矩形系は CSS で
        st = [base]
        if fill:
            st.append(f"background:{fill}")
        if line and lw:
            st.append(f"border:{lw:.1f}px {'dashed' if dash and 'dash' in dash else 'solid'} {line}")
        if prst in ("roundRect", "round1Rect", "round2SameRect", "snip1Rect"):
            st.append(f"border-radius:{min(w, h) * 0.16:.1f}px")
        elif prst in ("ellipse", "flowChartConnector"):
            st.append("border-radius:50%")
        elif prst == "can":
            st.append("border-radius:50% / 12%")
        shell_style = ";".join(st)

    txt = s.get("text")
    text_html = ""
    if txt:
        anchor = {"t": "flex-start", "ctr": "center", "b": "flex-end"}.get(txt.get("anchor") or "ctr",
                                                                          "center")
        vert = txt.get("vert")
        vcss = ""
        if vert in ("vert", "eaVert"):
            vcss = "writing-mode:vertical-rl;"
        elif vert == "vert270":
            vcss = "writing-mode:vertical-rl;transform:rotate(180deg);"
        paras = []
        for p in txt["paras"]:
            algn = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}.get(
                p.get("algn") or "ctr", "center")
            runs = []
            for run in p["runs"]:
                if run["text"] == "\n":
                    runs.append("<br>")
                    continue
                css = []
                if run.get("size"):
                    css.append(f'font-size:{run["size"]}pt')
                if run.get("b"):
                    css.append("font-weight:700")
                if run.get("i"):
                    css.append("font-style:italic")
                if run.get("u"):
                    css.append("text-decoration:underline")
                if run.get("color"):
                    css.append(f'color:{run["color"]}')
                if run.get("font"):
                    css.append(f'font-family:"{run["font"]}",{FONT_FALLBACK}')
                t = esc(run["text"])
                runs.append(f'<span style="{";".join(css)}">{t}</span>' if css else t)
            paras.append(f'<p style="text-align:{algn}">{"".join(runs)}</p>')
        text_html = (f'<div class="dtx" style="align-items:{anchor};{vcss}">'
                     f'{"".join(paras)}</div>')

    return f'<div class="dobj" style="{shell_style}">{"".join(body)}{text_html}</div>'


# ---------------------------------------------------------------------------
# 印刷設定（用紙・余白・改ページ・印刷範囲）
# ---------------------------------------------------------------------------

# Excel の用紙コード → (幅, 高さ) インチ
PAPER = {
    1: (8.5, 11), 2: (8.5, 11), 5: (8.5, 14), 7: (7.25, 10.5),
    8: (11.69, 16.54),    # A3
    9: (8.27, 11.69),     # A4
    11: (5.83, 8.27),     # A5
    12: (10.12, 14.33),   # B4 (JIS)
    13: (7.17, 10.12),    # B5 (JIS)
    43: (10.12, 14.33), 44: (7.17, 10.12),
}


def print_area_bounds(ws, opts) -> tuple[int, int, int, int] | None:
    """印刷範囲を (min_row, min_col, max_row, max_col) にする。複数指定は外接矩形。"""
    area = ws.print_area
    if not area:
        return None
    parts = area if isinstance(area, (list, tuple)) else [area]
    rows, cols = [], []
    for p in parts:
        for token in str(p).split(","):
            token = token.split("!")[-1].replace("$", "").strip()
            if not token:
                continue
            try:
                from openpyxl.utils import range_boundaries
                c1, r1, c2, r2 = range_boundaries(token)
            except Exception:
                continue
            if None in (c1, r1, c2, r2):
                continue
            rows += [r1, r2]
            cols += [c1, c2]
    if not rows:
        return None
    return (min(rows), min(cols), min(max(rows), opts.max_rows),
            min(max(cols), opts.max_cols))


def page_settings(ws, sheet_width_px: float) -> dict:
    """@page 用の用紙・余白・拡大率を組み立てる。"""
    ps, pm = ws.page_setup, ws.page_margins
    w_in, h_in = PAPER.get(ps.paperSize or 9, PAPER[9])
    landscape = (ps.orientation or "portrait") == "landscape"
    if landscape:
        w_in, h_in = h_in, w_in
    ml = pm.left if pm.left is not None else 0.7
    mr = pm.right if pm.right is not None else 0.7
    mt = pm.top if pm.top is not None else 0.75
    mb = pm.bottom if pm.bottom is not None else 0.75

    zoom = 1.0
    fit = getattr(ws.sheet_properties.pageSetUpPr, "fitToPage", False) \
        if ws.sheet_properties.pageSetUpPr is not None else False
    if fit:
        fw = ps.fitToWidth if ps.fitToWidth is not None else 1
        if fw:
            printable = (w_in - ml - mr) * 96.0
            if sheet_width_px > printable > 0:
                zoom = printable / sheet_width_px
    elif ps.scale:
        zoom = ps.scale / 100.0
    return {"w": w_in * 25.4, "h": h_in * 25.4, "m": (mt * 25.4, mr * 25.4, mb * 25.4, ml * 25.4),
            "zoom": zoom}


def title_row_range(ws) -> tuple[int, int] | None:
    t = ws.print_title_rows
    if not t:
        return None
    m = re.match(r"\$?(\d+):\$?(\d+)", str(t).split("!")[-1].replace("$", ""))
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


# ---------------------------------------------------------------------------
# 出力全体
# ---------------------------------------------------------------------------

PAGE_CSS = """
:root{color-scheme:light;--gl:%(grid)s}
body.nogl{--gl:transparent}
*{box-sizing:border-box}
body{margin:0;background:#f3f3f3;font-family:%(ff)s;color:#000}
.tabs{position:sticky;top:0;z-index:50;display:flex;gap:2px;padding:6px 8px 0;
  background:#f8f8f8;border-bottom:1px solid #d0d0d0;flex-wrap:wrap}
.tabs button{font:inherit;font-size:12px;padding:5px 14px;border:1px solid #cfcfcf;
  border-bottom:none;border-radius:5px 5px 0 0;background:#ececec;cursor:pointer;color:#333}
.tabs button.on{background:#fff;color:#107c41;font-weight:700;box-shadow:inset 0 2px 0 #107c41}
.tabs .tg{margin-left:auto;align-self:center;font-size:12px;color:#444;padding:0 6px 5px;
  display:flex;gap:5px;align-items:center;cursor:pointer;user-select:none}
.tabs .tg~.tg{margin-left:14px}   /* 2つ目以降は右端に並べる */
.wrap{padding:16px 48px 72px;overflow:auto}
.sheet{display:none}
.sheet.on{display:block}
.canvas{position:relative;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.18)}
table{border-collapse:collapse;table-layout:fixed;border-spacing:0;width:auto}
td{padding:0;position:relative;overflow:visible;vertical-align:bottom;
  font-size:11pt;line-height:1.18}
td.hidden{border:none!important;font-size:0}
td>i{font-style:normal;display:inline-block;z-index:1;padding:0 %(pad)spx;
  overflow:hidden;pointer-events:auto}
td>i.cr{display:flex;justify-content:flex-end}
/* 見切れているセル: クリックで全文表示を固定、もう一度クリックで元に戻す。
   切れている文字も DOM 上には残っているので、開かなくても選択コピー・検索はできる。 */
td>i.on{cursor:zoom-in}
td>i.on.open{cursor:zoom-out}
body.mark td>i.on:not(.open){box-shadow:inset -3px 0 0 rgba(198,74,60,.5)}
body.mark td>i.cr.on:not(.open){box-shadow:inset 3px 0 0 rgba(198,74,60,.5)}
body.mark td>i.cw.on:not(.open){box-shadow:inset 0 -3px 0 rgba(198,74,60,.5)}
td>i.on.open{overflow:visible;max-width:none!important;max-height:none!important;
  z-index:30;background:#fffdf0;box-shadow:0 0 0 1px #d9b95c,0 2px 8px rgba(0,0,0,.2)}
.measuring td>i{max-width:none!important;max-height:none!important}
/* 折り返しセルは絶対配置に切り替えて開く（行の高さを押し広げないため） */
td>i.cw.on.open{position:absolute;top:0;left:0;right:0;width:auto}
td a{color:#0563c1}
.draw{position:absolute;inset:0;pointer-events:none;z-index:5}
.dobj{position:absolute;pointer-events:auto}
img.dobj{object-fit:fill}
.dsvg{position:absolute;left:0;top:0;overflow:visible}
.dtx{position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center;
  padding:3px 5px;font-size:10.5pt;line-height:1.25}
.dtx p{margin:0}
@media print{
  :root{--gl:transparent}          /* Excel と同じく目盛線は印刷しない */
  body{background:#fff}
  .tabs{display:none}
  .wrap{padding:0;overflow:visible}
  .sheet{display:block!important;break-after:page}
  .sheet:last-child{break-after:auto}
  .canvas{box-shadow:none}
  tr.pb{break-before:page}          /* Excel の改ページ位置 */
  td>i.on{box-shadow:none}
}
"""


def build_html(title: str, sheets: list[dict]) -> str:
    css = [PAGE_CSS % {"ff": FONT_FALLBACK, "pad": CELL_PAD, "grid": GRID_COLOR}]
    for i, s in enumerate(sheets):
        # シートごとにクラスで名前空間を切る（id で切ると詳細度が高すぎて
        # 見切れ表示などの上書きが効かなくなる）
        scoped = []
        for line in s["css"].split("\n"):
            if line:
                scoped.append(f".s-{i} " + line)
        css.append("\n".join(scoped))
        pg = s.get("page")
        if pg:
            mt, mr, mb, ml = pg["m"]
            css.append(f"@page p{i}{{size:{pg['w']:.0f}mm {pg['h']:.0f}mm;"
                       f"margin:{mt:.0f}mm {mr:.0f}mm {mb:.0f}mm {ml:.0f}mm}}\n"
                       f".s-{i}{{page:p{i}}}")
            if abs(pg["zoom"] - 1.0) > 0.01:
                css.append(f"@media print{{.s-{i} .canvas{{zoom:{pg['zoom']:.4f}}}}}")
        if s.get("print_gridlines"):     # Excel 側で「枠線を印刷する」が有効なシート
            css.append(f"@media print{{.s-{i}{{--gl:{GRID_COLOR}}}}}")
    tabs = "".join(
        f'<button data-i="{i}" class="{"on" if i == 0 else ""}">{esc(s["name"])}</button>'
        for i, s in enumerate(sheets))
    bodies = "".join(
        f'<div class="sheet s-{i}{" on" if i == 0 else ""}" id="sh{i}">{s["html"]}</div>'
        for i, s in enumerate(sheets))
    script = """
function mark(sh){
  if(!sh||sh.dataset.marked||!sh.classList.contains('on'))return;
  sh.dataset.marked='1';
  // 右揃えの見切れは scrollWidth に出ないので、制限を外した幅と比べて判定する
  var cand=[].slice.call(sh.querySelectorAll('td>i.c,td>i.cr'));
  var lim=cand.map(function(el){return parseFloat(el.style.maxWidth)||0});
  sh.classList.add('measuring');
  var nat=cand.map(function(el){return el.getBoundingClientRect().width});
  sh.classList.remove('measuring');
  cand.forEach(function(el,i){if(lim[i]&&nat[i]>lim[i]+2)el.classList.add('on')});
  sh.querySelectorAll('td>i.cw').forEach(function(el){
    if(el.scrollHeight>el.clientHeight+2||el.scrollWidth>el.clientWidth+2)el.classList.add('on');
  });
}
document.querySelectorAll('.tabs button').forEach(function(b){
  b.addEventListener('click',function(){
    document.querySelectorAll('.tabs button').forEach(function(x){x.classList.remove('on')});
    document.querySelectorAll('.sheet').forEach(function(x){x.classList.remove('on')});
    b.classList.add('on');
    var sh=document.getElementById('sh'+b.dataset.i);
    sh.classList.add('on'); mark(sh);
  });
});
var cb=document.getElementById('markclip');
if(cb)cb.addEventListener('change',function(){
  document.body.classList.toggle('mark',cb.checked);
});
var gl=document.getElementById('gridlines');
if(gl)gl.addEventListener('change',function(){
  document.body.classList.toggle('nogl',!gl.checked);
});
// クリックで開閉。ドラッグ（範囲選択）と区別するため、押した位置から動いた場合と
// 文字が選択されている場合は無視する。
var px=0,py=0;
document.addEventListener('mousedown',function(e){px=e.clientX;py=e.clientY},true);
document.addEventListener('click',function(e){
  var el=e.target.closest?e.target.closest('td>i.on'):null;
  if(!el)return;
  if(Math.abs(e.clientX-px)>4||Math.abs(e.clientY-py)>4)return;
  var s=window.getSelection();
  if(s&&!s.isCollapsed)return;
  el.classList.toggle('open');
});
document.addEventListener('keydown',function(e){
  if(e.key==='Escape')document.querySelectorAll('td>i.open').forEach(function(el){
    el.classList.remove('open')});
});
mark(document.querySelector('.sheet.on'));
"""
    toggle = ('<label class="tg" title="Excel と同じ位置で文字が切れているセルに印を付けます。'
              '印の有無にかかわらず、クリックすると全文を表示できます">'
              '<input type="checkbox" id="markclip">見切れに印</label>')
    if any(s.get("gridlines") for s in sheets):
        toggle = ('<label class="tg" title="Excel の目盛線（薄いグレーの線）の表示を切り替えます。'
                  '印刷には元から出ません">'
                  '<input type="checkbox" id="gridlines" checked>目盛線</label>') + toggle
    tabs_html = f'<div class="tabs">{tabs if len(sheets) > 1 else ""}{toggle}</div>'
    return (f'<!doctype html>\n<html lang="ja"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1">'
            f"<title>{esc(title)}</title>\n<style>\n{''.join(css)}\n</style></head>\n"
            f'<body>{tabs_html}<div class="wrap">{bodies}</div>'
            f"<script>{script}</script></body></html>\n")


def check_container(src: str):
    """xlsx(zip) でないファイルを分かりやすく弾く。"""
    with open(src, "rb") as f:
        head = f.read(8)
    if head[:4] == b"\xd0\xcf\x11\xe0":
        raise SystemExit(
            "このファイルは xlsx ではありません（旧 .xls 形式、またはパスワード保護／IRM 付き）。\n"
            "Excel で開いて『名前を付けて保存 → Excel ブック(.xlsx)』で保存し直してから再実行してください。")
    if head[:2] != b"PK":
        raise SystemExit("xlsx として読めない形式です。")


def convert(src: str, dst: str, opts) -> str:
    check_container(src)
    wb = load_workbook(src, data_only=True, rich_text=True)
    resolver = ColorResolver(getattr(wb, "loaded_theme", None))
    parser = None
    if not opts.no_drawings:
        try:
            parser = DrawingParser(src, resolver)
        except Exception as e:      # 図形解析の失敗で全体を落とさない
            print(f"警告: 図形の解析に失敗しました ({e})", file=sys.stderr)
            parser = None

    targets = wb.sheetnames if not opts.sheet else [opts.sheet]
    formulas: dict[str, dict] = {}
    if opts.show_formula or opts.formula_tips or not opts.no_formula_check:
        try:
            formulas = load_formulas(src, targets)
        except Exception as e:
            print(f"警告: 数式の読み込みに失敗しました ({e})", file=sys.stderr)
    n_formula = sum(len(m) for m in formulas.values())
    missing_total = 0

    sheets = []
    for name in targets:
        if name not in wb.sheetnames:
            print(f"警告: シート '{name}' が見つかりません", file=sys.stderr)
            continue
        ws = wb[name]
        if ws.sheet_state != "visible" and not opts.hidden_sheets:
            continue
        if opts.gridlines == "on":
            gl = True
        elif opts.gridlines == "off":
            gl = False
        else:
            # 属性が省略されている場合、OOXML の既定は「表示する」
            sv = getattr(ws.sheet_view, "showGridLines", None)
            gl = True if sv is None else bool(sv)
        max_row = min(ws.max_row or 1, opts.max_rows)
        max_col = min(ws.max_column or 1, opts.max_cols)
        grid_for_shapes = Grid(ws, max_col, max_row)
        shapes = []
        if parser is not None:
            try:
                shapes = parser.shapes_for(name, grid_for_shapes)
            except Exception as e:
                print(f"警告: '{name}' の図形を読み込めませんでした ({e})", file=sys.stderr)
        bounds = print_area_bounds(ws, opts) if opts.print_area else None
        breaks = {b.id + 1 for b in ws.row_breaks.brk} if ws.row_breaks else set()
        res = render_sheet(ws, resolver, shapes, gl, name, opts, formulas.get(name),
                           bounds, breaks, title_row_range(ws))
        missing_total += res["missing"]
        page = page_settings(ws, res["width"]) if not opts.no_page_setup else None
        po = getattr(ws, "print_options", None)
        sheets.append({"name": name, "html": res["html"], "css": res["css"], "page": page,
                       "gridlines": gl,
                       "print_gridlines": bool(gl and po is not None and po.gridLines)})
        note = f", 計算結果なしの数式{res['missing']}個" if res["missing"] else ""
        area = "（印刷範囲で切り出し）" if bounds else ""
        print(f"  ・{name}: {res['cols']}列 × {res['rows']}行, 図形{len(shapes)}個{note}{area}")
    if parser is not None:
        parser.close()
    if not sheets:
        raise SystemExit("変換対象のシートがありません")

    if n_formula:
        print(f"  数式セル {n_formula} 個 → 保存済みの計算結果を表示"
              + (f"（うち {missing_total} 個は結果が未保存）" if missing_total else ""))
    if missing_total and not opts.show_formula:
        print("  ※ 計算結果が保存されていない数式セルは空欄になります。"
              "Excel で開いて上書き保存するか、--show-formula で数式そのものを出せます。",
              file=sys.stderr)

    title = os.path.splitext(os.path.basename(src))[0]
    out = build_html(title, sheets)
    with open(dst, "w", encoding="utf-8", newline="\n") as f:
        f.write(out)
    return dst


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Excel 方眼紙ドキュメントを見た目そのままの HTML に変換します。")
    ap.add_argument("input", help="入力 .xlsx / .xlsm")
    ap.add_argument("-o", "--output", help="出力 .html（既定: 入力と同名）")
    ap.add_argument("--sheet", help="変換するシート名（既定: 全シート）")
    ap.add_argument("--gridlines", choices=["auto", "on", "off"], default="auto",
                    help="目盛線の描画（既定 auto = Excel の設定に従う）")
    ap.add_argument("--no-drawings", action="store_true", help="画像・図形を出力しない")
    ap.add_argument("--hidden-sheets", action="store_true", help="非表示シートも出力する")
    ap.add_argument("--max-cols", type=int, default=1024, help="出力する最大列数")
    ap.add_argument("--max-rows", type=int, default=20000, help="出力する最大行数")
    ap.add_argument("--formula-tips", action="store_true",
                    help="数式セルにマウスを乗せると数式を表示する")
    ap.add_argument("--show-formula", action="store_true",
                    help="計算結果が保存されていない数式セルに、数式そのものを表示する")
    ap.add_argument("--no-formula-check", action="store_true",
                    help="数式の走査を省略して高速化する")
    ap.add_argument("--print-area", action="store_true",
                    help="印刷範囲が設定されているシートは、その範囲だけを出力する")
    ap.add_argument("--no-page-setup", action="store_true",
                    help="用紙サイズ・余白・改ページなどの印刷設定を反映しない")
    opts = ap.parse_args(argv)

    src = opts.input
    if not os.path.exists(src):
        raise SystemExit(f"入力が見つかりません: {src}")
    dst = opts.output or os.path.splitext(src)[0] + ".html"
    print(f"変換中: {src}")
    convert(src, dst, opts)
    size = os.path.getsize(dst)
    print(f"完了: {dst}  ({size/1024:.1f} KB)")


if __name__ == "__main__":
    main()
