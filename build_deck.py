#!/usr/bin/env python3
"""ساخت ارائه‌ی PowerPoint فارسی (راست‌به‌چپ) به کمک python-pptx.

محتوا در فایل content.py تعریف می‌شود و این فایل فقط ظاهر/چیدمان را می‌سازد.

نمونه اجرا:
    python3 build_deck.py                      # output/deck.pptx
    python3 build_deck.py -o output/v2.pptx    # خروجی دلخواه
    python3 build_deck.py --check              # فقط اعتبارسنجی فایل خروجی
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Inches, Pt

from content import DECK

# --------------------------------------------------------------------------
# تم و ابعاد
# --------------------------------------------------------------------------
W, H = 13.333, 7.5          # اینچ، نسبت ۱۶:۹
MARGIN = 0.85               # حاشیه‌ی چپ و راست
CONTENT_W = W - 2 * MARGIN

THEME = {
    "font": "Tahoma",       # فونت امن و موجود در Office؛ برای Vazirmatn مقدار را عوض کنید
    "ink": "0B1F33",
    "ink_soft": "1F3448",
    "body": "40505F",
    "muted": "7C8A97",
    "primary": "0F766E",
    "accent": "E8A33D",
    "surface": "F3F6F8",
    "line": "DBE2E8",
    "white": "FFFFFF",
}


def rgb(name: str) -> RGBColor:
    """رنگ را از کلید تم یا کد هگز می‌گیرد و به RGBColor تبدیل می‌کند."""
    return RGBColor.from_string(THEME.get(name, name).lstrip("#"))


# --------------------------------------------------------------------------
# کمک‌تابع‌های متنی (RTL)
# --------------------------------------------------------------------------
def _set_complex_script(run, typeface: str) -> None:
    """فونت خطوط فارسی/عربی در PowerPoint از a:cs خوانده می‌شود، نه a:latin."""
    rPr = run._r.get_or_add_rPr()
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = parse_xml(f'<a:cs {nsdecls("a")} typeface="{typeface}"/>')
        latin = rPr.find(qn("a:latin"))
        if latin is not None:
            latin.addnext(cs)
        else:
            ea = rPr.find(qn("a:ea"))
            if ea is not None:
                ea.addnext(cs)
            else:
                rPr.append(cs)
    cs.set("typeface", typeface)


def style_run(run, size: float, *, bold=False, color=None, font=None, italic=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font or THEME["font"]
    if color is not None:
        f.color.rgb = color
    _set_complex_script(run, font or THEME["font"])


def _apply_bullet(paragraph, indent_pt: float = 20) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    mar = str(Emu(int(Pt(indent_pt))))
    pPr.set("marL", mar)
    pPr.set("indent", str(-int(Pt(indent_pt))))
    bullet_font = parse_xml(f'<a:buFont {nsdecls("a")} typeface="Arial"/>')
    bullet_char = parse_xml(f'<a:buChar {nsdecls("a")} char="\u2022"/>')
    pPr.insert_element_before(
        bullet_font,
        "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip",
        "a:tabLst", "a:defRPr", "a:extLst",
    )
    pPr.insert_element_before(
        bullet_char, "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst"
    )


def _next_paragraph(text_frame):
    first = text_frame.paragraphs[0]
    if not first.runs and not first.text:
        return first
    return text_frame.add_paragraph()


def add_para(
    text_frame,
    text: str = "",
    *,
    size: float = 16,
    color=None,
    bold: bool = False,
    bullet: bool = False,
    space_before: float = 0,
    space_after: float = 6,
    line_spacing: float = 1.3,
    align=PP_ALIGN.RIGHT,
    font=None,
):
    """یک پاراگراف راست‌به‌چپ به انتهای کادر متن اضافه می‌کند."""
    p = _next_paragraph(text_frame)
    p.alignment = align
    p._p.get_or_add_pPr().set("rtl", "1")
    p.line_spacing = line_spacing
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    style_run(run, size, bold=bold, color=color or rgb("body"), font=font)
    if bullet:
        _apply_bullet(p)
    return p


def add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    anchor=MSO_ANCHOR.TOP,
    wrap: bool = True,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    return box


# --------------------------------------------------------------------------
# کمک‌تابع‌های گرافیکی
# --------------------------------------------------------------------------
def add_shape(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill=None,
    line=None,
    line_w: float = 1.0,
    shape=MSO_SHAPE.RECTANGLE,
    radius: float | None = None,
    name: str | None = None,
):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    if name:
        shp.name = name
    return shp


def set_background(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def to_fa(value) -> str:
    """ارقام لاتین را به فارسی تبدیل می‌کند."""
    return str(value).translate(FA_DIGITS)


# --------------------------------------------------------------------------
# اجزای تکرارشونده‌ی اسلاید
# --------------------------------------------------------------------------
def slide_header(slide, kicker, title) -> None:
    y = 0.66
    if kicker:
        box = add_textbox(slide, MARGIN, y, CONTENT_W, 0.32)
        add_para(
            box.text_frame, kicker, size=12.5, bold=True,
            color=rgb("primary"), space_after=0, line_spacing=1.0,
        )
        y += 0.34
    box = add_textbox(slide, MARGIN, y, CONTENT_W, 0.8)
    add_para(
        box.text_frame, title, size=30, bold=True,
        color=rgb("ink"), space_after=0, line_spacing=1.1,
    )
    add_shape(slide, W - MARGIN - 0.95, 2.02, 0.95, 0.055, fill=rgb("accent"))


def add_footer(slide, number: int, total: int, label: str) -> None:
    add_shape(slide, MARGIN, H - 0.86, CONTENT_W, 0.01, fill=rgb("line"))
    box = add_textbox(slide, MARGIN, H - 0.68, CONTENT_W - 1.6, 0.3)
    add_para(
        box.text_frame, label, size=9.5, color=rgb("muted"),
        space_after=0, line_spacing=1.0,
    )
    box = add_textbox(slide, W - MARGIN - 1.5, H - 0.68, 1.5, 0.3)
    add_para(
        box.text_frame, f"{to_fa(number)} / {to_fa(total)}", size=9.5,
        color=rgb("muted"), space_after=0, line_spacing=1.0,
        align=PP_ALIGN.LEFT,
    )


# --------------------------------------------------------------------------
# سازنده‌های هر نوع اسلاید
# --------------------------------------------------------------------------
def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build_hero(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("ink"))
    # نوار رنگی لبه‌ی راست + حلقه‌ی تزئینی
    add_shape(s, W - 0.16, 0, 0.16, H, fill=rgb("primary"))
    add_shape(
        s, -1.7, H - 3.0, 4.4, 4.4, shape=MSO_SHAPE.OVAL,
        line=rgb("ink_soft"), line_w=22,
    )

    y = 2.0
    if spec.get("kicker"):
        box = add_textbox(s, MARGIN, y, 10.4, 0.34)
        add_para(
            box.text_frame, spec["kicker"], size=13.5, bold=True,
            color=rgb("accent"), space_after=0, line_spacing=1.0,
        )
        y += 0.42
    box = add_textbox(s, MARGIN, y, 10.4, 1.7)
    add_para(
        box.text_frame, spec["title"], size=44, bold=True,
        color=rgb("white"), space_after=0, line_spacing=1.12,
    )
    if spec.get("subtitle"):
        box = add_textbox(s, MARGIN, y + 1.75, 9.2, 1.0)
        add_para(
            box.text_frame, spec["subtitle"], size=17.5,
            color=rgb("A9BACA"), space_after=0, line_spacing=1.45,
        )
    meta = spec.get("meta") or []
    if meta:
        box = add_textbox(s, MARGIN, H - 1.55, 9.2, 0.34)
        add_para(
            box.text_frame, "   |   ".join(str(m) for m in meta), size=12.5,
            color=rgb("muted"), space_after=0, line_spacing=1.0,
        )
    return s


def build_section(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("ink"))
    add_shape(s, W - 0.16, 0, 0.16, H, fill=rgb("primary"))
    if spec.get("index"):
        box = add_textbox(s, MARGIN, 1.55, CONTENT_W, 1.1)
        add_para(
            box.text_frame, to_fa(spec["index"]), size=64, bold=True,
            color=rgb("ink_soft"), space_after=0, line_spacing=1.0,
        )
    box = add_textbox(s, MARGIN, 2.85, 9.8, 1.3)
    add_para(
        box.text_frame, spec["title"], size=36, bold=True,
        color=rgb("white"), space_after=0, line_spacing=1.15,
    )
    add_shape(s, W - MARGIN - 0.95, 3.98, 0.95, 0.055, fill=rgb("accent"))
    if spec.get("subtitle"):
        box = add_textbox(s, MARGIN, 4.28, 9.4, 0.9)
        add_para(
            box.text_frame, spec["subtitle"], size=15.5,
            color=rgb("A9BACA"), space_after=0, line_spacing=1.45,
        )
    return s


def build_agenda(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("surface"))
    slide_header(s, spec.get("kicker"), spec["title"])

    items = spec.get("items", [])
    top, bottom = 2.42, 6.35
    step = min(0.72, (bottom - top) / max(len(items), 1))
    for i, item in enumerate(items):
        y = top + i * step
        badge = add_shape(
            s, W - MARGIN - 0.46, y + 0.04, 0.46, 0.46,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22,
            fill=rgb("primary"),
        )
        tf = badge.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        add_para(
            tf, to_fa(i + 1), size=12.5, bold=True, color=rgb("white"),
            space_after=0, line_spacing=1.0, align=PP_ALIGN.CENTER,
        )
        box = add_textbox(s, MARGIN, y, CONTENT_W - 0.76, step - 0.08)
        add_para(
            box.text_frame, item, size=17.5, color=rgb("ink"),
            space_after=0, line_spacing=1.2,
        )
        if i < len(items) - 1:
            add_shape(
                s, MARGIN, y + step - 0.13, CONTENT_W - 0.76, 0.008,
                fill=rgb("line"),
            )
    add_footer(s, ctx["number"], ctx["total"], ctx["label"])
    return s


def build_bullets(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("surface"))
    slide_header(s, spec.get("kicker"), spec["title"])

    items = spec.get("items", [])
    size = spec.get("size", 18 if len(items) <= 5 else 16)
    box = add_textbox(s, MARGIN, 2.36, CONTENT_W - 0.1, 3.9)
    for item in items:
        add_para(
            box.text_frame, item, size=size, color=rgb("body"), bullet=True,
            space_after=spec.get("space_after", 16), line_spacing=1.5,
        )

    note = spec.get("note")
    if note:
        add_shape(s, W - MARGIN - 0.045, 6.28, 0.045, 0.46, fill=rgb("accent"))
        box = add_textbox(s, MARGIN, 6.3, CONTENT_W - 0.28, 0.5)
        add_para(
            box.text_frame, note, size=12.5, color=rgb("primary"),
            bold=True, space_after=0, line_spacing=1.2,
        )
    add_footer(s, ctx["number"], ctx["total"], ctx["label"])
    return s



def _grid(slide, items, *, top, height, builder):
    """آیتم‌ها را به ترتیب از راست به چپ در یک ردیف می‌چیند."""
    n = max(len(items), 1)
    gap = 0.26
    width = (CONTENT_W - gap * (n - 1)) / n
    for i, item in enumerate(items):
        x = W - MARGIN - width - i * (width + gap)
        builder(slide, item, i, x, top, width, height)


def build_cards(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("surface"))
    slide_header(s, spec.get("kicker"), spec["title"])

    def card(slide, item, i, x, y, w, h):
        add_shape(
            slide, x, y, w, h, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06,
            fill=rgb("white"), line=rgb("line"), line_w=1,
        )
        add_shape(slide, x + w - 0.34, y + 0.3, 0.18, 0.05, fill=rgb("accent"))
        box = add_textbox(slide, x + 0.28, y + 0.28, w - 0.6, 0.3)
        add_para(
            box.text_frame, to_fa(f"{i + 1:02d}"), size=11.5, bold=True,
            color=rgb("muted"), space_after=0, line_spacing=1.0,
        )
        box = add_textbox(slide, x + 0.28, y + 0.66, w - 0.56, 0.6)
        add_para(
            box.text_frame, item.get("title", ""), size=16.5, bold=True,
            color=rgb("ink"), space_after=0, line_spacing=1.2,
        )
        body = add_textbox(slide, x + 0.28, y + 1.34, w - 0.56, h - 1.6)
        add_para(
            body.text_frame, item.get("body", ""), size=13,
            color=rgb("body"), space_after=0, line_spacing=1.55,
        )

    _grid(s, spec.get("items", []), top=2.42, height=3.5, builder=card)
    add_footer(s, ctx["number"], ctx["total"], ctx["label"])
    return s


def build_stats(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("surface"))
    slide_header(s, spec.get("kicker"), spec["title"])

    def stat(slide, item, i, x, y, w, h):
        add_shape(
            slide, x, y, w, 1.95, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
            radius=0.08, fill=rgb("white"), line=rgb("line"), line_w=1,
        )
        box = add_textbox(slide, x, y + 0.42, w, 0.72)
        add_para(
            box.text_frame, item.get("value", ""), size=38, bold=True,
            color=rgb("primary"), space_after=0, line_spacing=1.0,
            align=PP_ALIGN.CENTER,
        )
        box = add_textbox(slide, x + 0.15, y + 1.2, w - 0.3, 0.62)
        add_para(
            box.text_frame, item.get("label", ""), size=12.5,
            color=rgb("muted"), space_after=0, line_spacing=1.3,
            align=PP_ALIGN.CENTER,
        )

    _grid(s, spec.get("items", []), top=2.5, height=1.95, builder=stat)
    body = spec.get("body")
    if body:
        box = add_textbox(s, MARGIN, 4.95, CONTENT_W, 1.2)
        add_para(
            box.text_frame, body, size=14.5, color=rgb("body"),
            space_after=0, line_spacing=1.6,
        )
    add_footer(s, ctx["number"], ctx["total"], ctx["label"])
    return s


def build_closing(prs, spec, ctx):
    s = new_slide(prs)
    set_background(s, rgb("ink"))
    add_shape(s, W - 0.16, 0, 0.16, H, fill=rgb("primary"))
    add_shape(
        s, W - 3.2, -1.6, 4.6, 4.6, shape=MSO_SHAPE.OVAL,
        line=rgb("ink_soft"), line_w=22,
    )
    box = add_textbox(s, MARGIN, 2.55, 9.8, 1.2)
    add_para(
        box.text_frame, spec["title"], size=42, bold=True,
        color=rgb("white"), space_after=0, line_spacing=1.1,
    )
    add_shape(s, W - MARGIN - 0.95, 3.78, 0.95, 0.055, fill=rgb("accent"))
    if spec.get("subtitle"):
        box = add_textbox(s, MARGIN, 4.1, 9.2, 0.9)
        add_para(
            box.text_frame, spec["subtitle"], size=16,
            color=rgb("A9BACA"), space_after=0, line_spacing=1.45,
        )
    meta = spec.get("meta") or []
    if meta:
        box = add_textbox(s, MARGIN, H - 1.55, 9.4, 0.34)
        add_para(
            box.text_frame, "   |   ".join(str(m) for m in meta), size=12.5,
            color=rgb("accent"), space_after=0, line_spacing=1.0,
        )
    return s


BUILDERS = {
    "hero": build_hero,
    "section": build_section,
    "agenda": build_agenda,
    "bullets": build_bullets,
    "cards": build_cards,
    "stats": build_stats,
    "closing": build_closing,
}




# --------------------------------------------------------------------------
# ساخت فایل و اعتبارسنجی
# --------------------------------------------------------------------------
def build(deck: dict, out_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    slides = deck["slides"]
    total = len(slides)
    label = deck["meta"].get("footer", deck["meta"].get("title", ""))

    unknown = [s.get("type") for s in slides if s.get("type") not in BUILDERS]
    if unknown:
        raise ValueError(f"نوع اسلاید ناشناخته: {unknown}")

    for idx, spec in enumerate(slides, start=1):
        ctx = {"number": idx, "total": total, "label": label}
        BUILDERS[spec["type"]](prs, spec, ctx)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path


def check(path: Path) -> bool:
    """فایل خروجی را دوباره باز می‌کند و خلاصه‌ی محتوا را چاپ می‌کند."""
    prs = Presentation(path)
    ok = True
    print(f"فایل: {path}  ({path.stat().st_size / 1024:.1f} KB)")
    print(f"ابعاد: {prs.slide_width.inches:.2f} x {prs.slide_height.inches:.2f} اینچ")
    print(f"تعداد اسلاید: {len(prs.slides)}")
    for i, slide in enumerate(prs.slides, start=1):
        texts = [
            sh.text_frame.text.strip().replace("\n", " / ")
            for sh in slide.shapes
            if sh.has_text_frame and sh.text_frame.text.strip()
        ]
        if not texts:
            ok = False
            print(f"  [{i}] خالی!")
            continue
        print(f"  [{i}] {texts[0][:70]}")
    return ok


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="ساخت ارائه‌ی فارسی")
    parser.add_argument(
        "-o", "--out", default="output/deck.pptx", help="مسیر فایل خروجی"
    )
    parser.add_argument(
        "--check", action="store_true", help="پس از ساخت، فایل را بررسی کن"
    )
    args = parser.parse_args(argv)

    out = Path(args.out)
    build(DECK, out)
    print(f"ساخته شد: {out}")
    if args.check:
        return 0 if check(out) else 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
