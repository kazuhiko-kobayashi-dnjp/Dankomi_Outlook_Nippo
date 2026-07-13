#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excelセル内の部分書式(太字/下線/文字色などのリッチテキスト)と、
Web側で使っているHTML文字列(contenteditable由来)を相互変換するブリッジ。

- celltext_to_html(): openpyxlのセル値(str または CellRichText) → HTML文字列
- html_to_richtext(): HTML文字列 → openpyxlに書き込み可能な値(str または CellRichText)
"""
import html
from html.parser import HTMLParser

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def _escape(s):
    return html.escape(s, quote=True)


def celltext_to_html(value, font=None):
    """openpyxlのセル値をWeb表示用のHTML文字列に変換する。

    value が CellRichText(部分書式あり)の場合はそのランごとの書式を反映する。
    value が素の文字列の場合でも、font(そのセル自体のFontオブジェクト)が渡されれば
    セル全体に適用された太字/下線/文字色をHTMLに反映する
    (Excelでは「セル全体を選択して太字」にした場合、CellRichTextにはならず
    通常の文字列値+セルのフォント設定という形になるため、これを別途拾う必要がある)。
    """
    if value is None:
        return ""
    if isinstance(value, CellRichText):
        parts = []
        for item in value:
            if isinstance(item, TextBlock):
                text = _escape(item.text).replace("\n", "<br>")
                font = item.font
                if font and font.b:
                    text = f"<b>{text}</b>"
                if font and font.u:
                    text = f"<u>{text}</u>"
                color = getattr(font.color, "rgb", None) if (font and font.color) else None
                if isinstance(color, str) and len(color) >= 6 and color[-6:].upper() != "000000":
                    hexcolor = color[-6:]
                    text = f'<font color="#{hexcolor}">{text}</font>'
                parts.append(text)
            else:
                parts.append(_escape(str(item)).replace("\n", "<br>"))
        return "".join(parts)
    text = _escape(str(value)).replace("\n", "<br>")
    if font:
        if font.b:
            text = f"<b>{text}</b>"
        if font.u:
            text = f"<u>{text}</u>"
        color = getattr(font.color, "rgb", None) if font.color else None
        if isinstance(color, str) and len(color) >= 6 and color[-6:].upper() != "000000":
            hexcolor = color[-6:]
            text = f'<font color="#{hexcolor}">{text}</font>'
    return text


class _RichTextHTMLParser(HTMLParser):
    """b/u/font color/br/div 程度の限定タグを解釈してフラット化する簡易パーサー。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.segments = []  # (text, bold, underline, color)
        self._bold = 0
        self._underline = 0
        self._color_stack = []
        self._seen_block = False

    def handle_starttag(self, tag, attrs):
        if tag == "b" or tag == "strong":
            self._bold += 1
        elif tag == "u":
            self._underline += 1
        elif tag == "font":
            color = dict(attrs).get("color")
            self._color_stack.append(color)
        elif tag == "br":
            self._append("\n")
        elif tag in ("div", "p"):
            if self._seen_block:
                self._append("\n")
            self._seen_block = True

    def handle_endtag(self, tag):
        if tag in ("b", "strong"):
            self._bold = max(0, self._bold - 1)
        elif tag == "u":
            self._underline = max(0, self._underline - 1)
        elif tag == "font":
            if self._color_stack:
                self._color_stack.pop()

    def handle_data(self, data):
        self._append(data)

    def _append(self, text):
        if not text:
            return
        color = self._color_stack[-1] if self._color_stack else None
        self.segments.append((text, bool(self._bold), bool(self._underline), color))


def html_to_richtext(html_str):
    """HTML文字列をExcelセルに書き込める値(素の文字列 or CellRichText)に変換する。"""
    if not html_str:
        return None
    if "<" not in html_str:
        return html_str

    parser = _RichTextHTMLParser()
    parser.feed(html_str)
    if not parser.segments:
        return None

    merged = []
    for text, bold, underline, color in parser.segments:
        if merged and merged[-1][1:] == (bold, underline, color):
            merged[-1] = (merged[-1][0] + text, bold, underline, color)
        else:
            merged.append((text, bold, underline, color))

    if all(not b and not u and not c for _, b, u, c in merged):
        return "".join(t for t, _, _, _ in merged) or None

    items = []
    for text, bold, underline, color in merged:
        if not bold and not underline and not color:
            items.append(text)
            continue
        font_kwargs = {}
        if bold:
            font_kwargs["b"] = True
        if underline:
            font_kwargs["u"] = "single"
        if color:
            hexcolor = color.lstrip("#").upper()
            if len(hexcolor) == 6:
                hexcolor = "FF" + hexcolor
            font_kwargs["color"] = hexcolor
        items.append(TextBlock(InlineFont(**font_kwargs), text))
    return CellRichText(*items) if items else None
