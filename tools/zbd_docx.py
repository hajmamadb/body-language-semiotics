# -*- coding: utf-8 -*-
"""
zbd_docx.py — سازندۀ اسکناس/مقالهٔ فارسی (RTL) با python-docx
==============================================================
ساخت اسکلت کامل مقالۀ فارسی با:
- جهت راست‌به‌چپ برای همهٔ پراگراف‌ها و جداول
- قلم فارسی (B Nazanin با فالبک Tahoma)
- بخش‌های استاندارد مقالهٔ علمی-پژوهشی
- جاهای خالی (placeholder) برای پژوهشگر: [درج ...]
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

FA = "B Nazanin"
FA_FALLBACK = "Tahoma"


def _set_fa_run(run, size: int = 12, bold: bool = False, color=None) -> None:
    run.font.name = FA
    run.font.size = Pt(size)
    run.bold = bold
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FA)
    rFonts.set(qn("w:hAnsi"), FA)
    rFonts.set(qn("w:cs"), FA)
    if color:
        run.font.color.rgb = color


def _rtl(par) -> None:
    pPr = par._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = pPr.makeelement(qn("w:bidi"), {})
        pPr.append(bidi)
    bidi.set(qn("w:val"), "1")
    # راست‌چین
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _add_par(doc, text, size=12, bold=False, color=None, space_after=6, indent_first=True):
    par = doc.add_paragraph()
    _rtl(par)
    run = par.add_run(text)
    _set_fa_run(run, size=size, bold=bold, color=color)
    par.paragraph_format.space_after = Pt(space_after)
    if indent_first:
        par.paragraph_format.first_line_indent = Pt(24)
    return par


def _add_heading(doc, text, level=1):
    sizes = {1: 16, 2: 14, 3: 12.5}
    par = doc.add_paragraph()
    _rtl(par)
    run = par.add_run(text)
    _set_fa_run(run, size=sizes.get(level, 12), bold=True,
                color=RGBColor(0x1F, 0x3A, 0x5F))
    par.paragraph_format.space_before = Pt(14)
    par.paragraph_format.space_after = Pt(6)
    par.paragraph_format.first_line_indent = Pt(0)
    return par


def _add_table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.allow_autofit = False
    # RTL table
    tbl = t._tbl
    tblPr = tbl.tblPr
    bidiVisual = tblPr.makeelement(qn("w:bidiVisual"), {})
    tblPr.append(bidiVisual)
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        par = hdr[i].paragraphs[0]
        _rtl(par)
        run = par.add_run(h)
        _set_fa_run(run, size=10.5, bold=True)
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            par = cells[i].paragraphs[0]
            _rtl(par)
            run = par.add_run(str(v))
            _set_fa_run(run, size=10.5)
            par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if widths:
        from docx.shared import Cm
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t


# ------------------------------------------------------------------
def build_article_docx(article: int, out: Path) -> Path:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    # پیش‌فند سراسری
    style = doc.styles["Normal"]
    style.font.name = FA
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:cs"), FA)

    if article == 1:
        title = ("نشانه‌شناسی ساختار کالبدی و زبان بدن سیدعلی خامنه‌ای در "
                 "موقعیت‌های ارتباطی منتخب: رویکرد حرکت‌شناختی و فرانقش‌های "
                 "دیداری خودبسنده")
        _add_par(doc, title, size=18, bold=True, space_after=4, indent_first=False)
        _add_par(doc, "(اسکلت پژوهش — جاهای [درج ...] را کامل کنید)",
                 size=10, color=RGBColor(0x80, 0x80, 0x80), indent_first=False, space_after=12)

        _add_heading(doc, "چکیده", 1)
        _add_par(doc, "[درج چکیده: هدف، روش (نشانه‌شناسی اجتماعی دیداری + حرکت‌شناسی "
                      "ساختاری/Kinesics)، داده (فیلم‌های ویدیویی از پنج موقعیت ارتباطی منتخب، "
                      "صامت‌سازی‌شده پیش از تحلیل)، واحد تحلیل (کینم‌های بدنی در فریم‌های "
                      "کلیدی و برش‌های پیکربندی)، یافته‌های اصلی و واژگان کلیدی]. "
                      "حدود ۲۰۰–۲۵۰ واژه.")
        _add_par(doc, "واژگان کلیدی: نشانه‌شناسی اجتماعی دیداری؛ Kinesics (حرکت‌شناسی "
                      "ساختاری)؛ فرانقش‌های بصری (Visual Metafunctions)؛ زبان بدن خودبسنده؛ "
                      "بردار نگاه (Gaze Vector)؛ هم‌نشینی نشانه‌ها؛ سیدعلی خامنه‌ای.", size=11)

        _add_heading(doc, "۱. مقدمه و بیان مسئله", 1)
        _add_par(doc, "[درج مقدمه: اهمیت زبان بدن در تولید معنا؛ مسئلهٔ پژوهش: چگونه "
                      "زبان بدن به‌مثابهٔ متن دیداری خودبسنده، فارغ از میانجی‌گری گفتار، "
                      "معنا می‌سازد؟ سؤالات فرعی: (الف) الگوهای هم‌نشینی نشانه‌های بدنی "
                      "(نگاه، سر، تنه، دست‌ها، پوشش) چگونه در نبود گفتار، فرانقش‌های دیداری "
                      "مفرد/بین‌شخصی/ترکیبی را برمی‌انگیزند؟ (ب) بردار نگاه (کنش Demand در "
                      "مقابل Offer) و انحراف محوری سر در هر موقعیت چه الگویی دارد؟ (ج) دامنهٔ "
                      "فضایی دست کنش‌گر (چپ) و زاویهٔ ستون فقرات چگونه روابط هژمونیک را بازتعریف "
                      "می‌کند؟ (د) مرزبندی پوشش (به‌ویژه جمع‌شدگی/گشودگی عبا) چگونه "
                      "Framing کالبدی می‌سازد؟ (ه) تفاوت‌های ساختاری میان موقعیت‌های "
                      "رسمی-هژمونیک (دیپلماتیک/نظامی) و غیررسمی-عاطفی (خانواده‌های شهدا/"
                      "شنیداری محض) چیست؟ تأکید بر مرز پژوهش: تفسیر نشانه‌های قابل‌مشاهده، "
                      "نه استنباط شخصیت/نیت.]")

        _add_heading(doc, "۲. چارچوب نظری", 1)
        _add_par(doc, "[درج چارچوب: نشانه‌شناسی سوسور/پیرس به‌عنوان پشتوانه؛ نشانه‌شناسی "
                      "اجتماعی دیداری (کرس و ون‌لیوون ۲۰۰۶) به‌عنوان چارچوب اصلی؛ حرکت‌شناسی "
                      "ساختاری Kinesics (Birdwhistell 1970) برای تعریف کینم؛ تقسیم‌بندی "
                      "Ekman & Friesen (1969) برای نشانه‌های بدنی؛ Alkingsen (1984) و "
                      "Stewart et al. (2015) برای تحلیل حرکتی. تمایز نشانهٔ بدنیِ خودبسنده از "
                      "نشانهٔ هم‌زمان با گفتار (co-speech).]")
        _add_par(doc, "اصل روش‌شناختی کلیدی: دلالت نشانهٔ بدنی نه در خود یک نشانهٔ مجزا، "
                      "بلکه در پیکربندی (Configuration) — هم‌نشینی چند کد در یک فریم — شکل "
                      "می‌گیرد. صدا پیش از تحلیل صامت‌سازی می‌شود تا ریتم، لحن و مفاهیم سخن "
                      "موجب سوگیری نشود.")

        _add_heading(doc, "۳. روش‌شناسی", 1)
        _add_par(doc, "روش: نشانه‌شناسی اجتماعی دیداری + حرکت‌شناسی ساختاری؛ تحلیل "
                      "محتوای فریم‌های منتخب با کدگذاری توصیفی ماتریسی (هر مقوله در یک یا "
                      "چند زیرمحور هندسی).")
        _add_heading(doc, "۳.۱. جامعه و نمونهٔ داده", 2)
        _add_par(doc, "داده: فیلم‌های نسبتا‌پیوسته و کم‌تدوین از پنج موقعیت ارتباطی "
                      "تمایز‌یافته:")
        from zbd_common import CONTEXTS as _CTX
        for _k, _c in _CTX.items():
            _add_par(doc, f"• {_c}", indent_first=False)
        _add_heading(doc, "۳.۲. واحد تحلیل", 2)
        _add_par(doc, "واحد تحلیل: «کینم‌های بدنی» در فریم‌های کلیدی و برش‌های "
                      "پیکربندی (Configuration Cuts) درون توالی‌های ۲–۳ دقیقه‌ای. تغییر "
                      "معنادار در ترکیب «نگاه + زاویهٔ سر + وضعیت دست»، پایان یک واحد و "
                      "آغاز واحد جدید را رقم می‌زند. شناسهٔ واحد به‌صورت "
                      "SIT-XXX-YYYYMM-FNN (موقعیت-کد-تاریخ-فریم) ثبت می‌شود.")
        _add_heading(doc, "۳.۳. ابزار و کتاب‌کد ماتریسی", 2)
        _add_par(doc, "کتاب‌کد ماتریسی هفت‌مقوله‌ای (کدهای G/F/H/T/L/R/C) — پیوست ۱. "
                      "هر مقوله در یک یا چند زیرمحور هندسی کدگذاری می‌شود؛ کد صرفاً توصیف "
                      "نشانهٔ قابل‌مشاهده است و دلالت (شفقت نمادین، صلابت نهادی، اقتدار "
                      "مافوق) در ستون‌های «محور دلالت»، «توضیح» و «ماتریس هم‌نشینی» (پیوست ۳) "
                      "نوشته می‌شود.")
        _add_heading(doc, "۳.۴. قواعد کدگذاری (راهنمای روش‌شناختی)", 2)
        from zbd_common import CODING_GUIDELINES as _GL
        for _t, _b in _GL:
            _add_par(doc, f"{_t}: {_b}", indent_first=False)
        _add_heading(doc, "۳.۵. اعتبار و پایایی", 2)
        _add_par(doc, "کدگذاری توسط دو ارزیاب مستقل انجام می‌شود و مقدار توافق بر مبنای "
                      "ضریب کاپای کوهن (Cohen's Kappa) سنجیده می‌شود؛ تنها کدهایی با "
                      "کاپای ≥ ۰٫۷۵ وارد تحلیل نهایی می‌شوند. محدودیت‌های قاب دوربین، "
                      "زاویهٔ نامناسب، کلوزآپ بیش از حد و تدوین در ستون «محدودیت تصویر» ثبت "
                      "و در تفسیر لحاظ می‌شوند.")

        _add_heading(doc, "۴. یافته‌ها", 1)
        _add_par(doc, "۴.۱. تواتر کدهای هر مقوله به‌جز بَعد [درج جدول تواتر + نمودار]")
        _add_par(doc, "۴.۲. دلالت پیکربندی‌ها در بافت هر موقعیت [درج نمونه‌های تحلیل‌شده]")
        _add_par(doc, "۴.۳. ماتریس هم‌نشینی: پیکربندی‌های الگویی و دلالت برآمده "
                      "[درج + مقایسه با پیوست ۳]")
        _add_par(doc, "۴.۴. تفاوت‌های ساختاری میان موقعیت‌های رسمی-هژمونیک و "
                      "غیررسمی-عاطفی [درج]")

        _add_heading(doc, "۵. بحث", 1)
        _add_par(doc, "[درج: تبیین یافته‌ها در پرتو چارچوب نشانه‌شناسی اجتماعی دیداری و "
                      "Kinesics؛ مقایسه با ادبیات (Ekman & Friesen, Birdwhistell, Alkingsen, "
                      "Stewart et al., Kress & van Leeuwen)؛ تأکید بر اینکه دلالت "
                      "پیکربندی‌محور است و تعمیم‌پذیری محدود به بافت‌های مشابه است.]")

        _add_heading(doc, "۶. نتیجه‌گیری", 1)
        _add_par(doc, "[درج: پاسخ مستقیم به سؤال اصلی و سؤالات فرعی؛ محدودیت‌ها و پیشنهادهای "
                      "پژوهش‌های آتی (شامل پیوند با مقالۀ دوم دربارهٔ بازنمایی رسانه‌ای).]")

        _add_heading(doc, "پیوست ۱: کتاب‌کد ماتریسی نشانه‌های کالبدی", 1)
        _add_table(doc,
                   ["مقوله", "کد", "عنوان (توصیف هندسی/ساختاری)"],
                   _codebook_rows(),
                   widths=[4, 1.5, 10])
        _add_heading(doc, "پیوست ۲: جدول کدگذاری نمونه", 1)
        _add_par(doc, "[درج چند سطر از برگهٔ کدگذاری به‌عنوان نمونهٔ شفاف‌سازی؛ شناسهٔ "
                      "واحد به‌صورت SIT-XXX-YYYYMM-FNN]")
        _add_heading(doc, "پیوست ۳: ماتریس هم‌نشینی نشانه‌ها", 1)
        _add_table(doc,
                   ["موقعیت", "پیکربندی (هم‌نشینی کدها)", "کدها", "دلالت برآمده"],
                   _matrix_rows(),
                   widths=[3.5, 6, 3.5, 4.5])

    else:  # article 2
        title = ("نشانه‌شناسی بازنمایی زبان بدن امام شهید در تصاویر صفحهٔ اول "
                 "روزنامه‌های منتخب: مطالعهٔ تطبیقی رویدادهای مشترک")
        _add_par(doc, title, size=18, bold=True, space_after=4, indent_first=False)
        _add_par(doc, "(اسکلت پژوهش — جاهای [درج ...] را کامل کنید)",
                 size=10, color=RGBColor(0x80, 0x80, 0x80), indent_first=False, space_after=12)

        _add_heading(doc, "چکیده", 1)
        _add_par(doc, "[درج چکیده: هدف، روش (نشانه‌شناسی اجتماعی + تحلیل تطبیقی بازنمایی)، "
                      "داده (صفحهٔ اول روزنامه‌های منتخب پس از رویدادهای مشترک)، یافته‌ها، "
                      "واژگان کلیدی].")
        _add_par(doc, "واژگان کلیدی: بازنمایی؛ زبان بدن؛ صفحهٔ اول روزنامه؛ "
                      "نشانه‌شناسی اجتماعی؛ [واژگان تکمیلی].", size=11)

        _add_heading(doc, "۱. مقدمه و بیان مسئله", 1)
        _add_par(doc, "[درج: تصاویر صفحهٔ اول صرفاً ثبت رویداد نیستند؛ انتخاب لحظه، قاب‌بندی، "
                      "برش، اندازه و جایگاه و نسبت با تیتر و شرح، معنا را می‌سازند. مسئله: "
                      "چگونگی بازنمایی و معناپردازی زبان بدن از طریق عکس‌ها و نسبت آن‌ها با "
                      "عناصر کلامی/بصری صفحه.]")

        _add_heading(doc, "۲. چارچوب نظری", 1)
        _add_par(doc, "[درج: نشانه‌شناسی اجتماعی؛ نظریه‌های بازنمایی رسانه‌ای (مثلاً "
                      "فیرکل/من‌وز) در حد ضرورت؛ تمایز «حالت بدنیِ ثبت‌شده» از «مسیر حرکت».]")
        _add_par(doc, "اصل روش‌شناختی: از یک عکس ثابت، مسیر یا تداوم حرکت استنباط نمی‌شود؛ "
                      "موضوع، حالت بدنیِ ثبت‌شده در آن لحظه است.")

        _add_heading(doc, "۳. روش‌شناسی", 1)
        _add_par(doc, "روش: نشانه‌شناسی اجتماعی با مقایسهٔ تطبیقی بازنمایی‌ها در پوشش یک "
                      "رویداد مشترک توسط چند روزنامه.")
        _add_heading(doc, "۳.۱. داده", 2)
        _add_par(doc, "داده: نسخهٔ کامل صفحهٔ اول روزنامه‌های منتخب در بازهٔ زمانی پس از "
                      "رویدادهای مشترک. عکس، کانون تحلیل است ولی در ارتباط با تیتر، شرح و "
                      "آرایش صفحه خوانده می‌شود.")
        _add_heading(doc, "۳.۲. واحد تحلیل و کتاب‌کد", 2)
        _add_par(doc, "واحد تحلیل: هر تصویر (به‌همراه عناصر همراه در صفحه). کدگذاری: نشانهٔ "
                      "بدنی ثبت‌شده + قاب‌بندی/زاویه/اندازه/برش + اندازه/جایگاه + نسبت با تیتر "
                      "و شرح + آرایهٔ مجاور + محور دلالت بازنمایی.")
        _add_heading(doc, "۳.۳. اعتبار و محدودیت", 2)
        _add_par(doc, "[درج: پایایی بین‌بازکننده؛ محدودیت: «انتخاب تصویر» قابل مطالعه است اما "
                      "توضیح نیت سردبیر بدون شواهد تکمیلی (مصاحبه، آرشیو) ممکن نیست.]")

        _add_heading(doc, "۴. یافته‌ها", 1)
        _add_par(doc, "۴.۱. نشانه‌های بدنیِ برجسته‌شده در تصاویر [درج]")
        _add_par(doc, "۴.۲. مشارکت قاب‌بندی، زاویه، اندازه و برش در دلالت‌پردازی [درج]")
        _add_par(doc, "۴.۳. نقش اندازه/جایگاه و نسبت با تیتر/شرح/آرایه [درج]")
        _add_par(doc, "۴.۴. شباهت و تفاوت بازنمایی میان روزنامه‌ها در پوشش رویداد مشترک [درج]")

        _add_heading(doc, "۵. بحث", 1)
        _add_par(doc, "[درج: تبیین بازنمایی‌ها؛ آیا تفاوت‌ها صرفاً در انتخاب لحظه است یا در "
                      "آرایش صفحه؟ قیود استنتاج دربارهٔ نیت رسانه.]")

        _add_heading(doc, "۶. نتیجه‌گیری", 1)
        _add_par(doc, "[درج: پاسخ به سؤالات؛ محدودیت‌ها؛ پیوند تطبیقی با مقالۀ اول (نشان بدنی "
                      "در بافت وقوع در برابر بازنمایی آن در بافت نشر).]")

        _add_heading(doc, "پیوست ۱: کتاب‌کد بازنمایی", 1)
        _add_table(doc,
                   ["بَعد", "کد/گزینه", "عنوان"],
                   _repr_codebook_rows(), widths=[3, 3, 11])
        _add_heading(doc, "پیوست ۲: شیت مقایسهٔ نمونه (رویداد مشترک)", 1)
        _add_par(doc, "[درج تصویر شیت مقایسهٔ ۳ صفحهٔ اول با منبع]")

    doc.save(out)
    return out


def _codebook_rows():
    from zbd_common import CODEBOOK
    titles = {
        "gaze": "بردار نگاه (Gaze Vector)",
        "face": "ریزآرایش چهره (Facial Micro-expression)",
        "head": "انحراف محوری سر (Head Orientation)",
        "torso": "راستای ستون فقرات (Torso Orientation)",
        "left_hand": "دامنهٔ فضایی دست چپ — دست کنش‌گر (Agent Hand)",
        "right_hand": "وضعیت دست راست — دست تکیه‌گاه (Support Hand)",
        "clothing": "مرزبندی پوشش (Framing)",
    }
    rows = []
    for dim, items in CODEBOOK.items():
        for c, l in items:
            rows.append((titles.get(dim, dim), c, l))
    return rows


def _matrix_rows():
    from zbd_common import MATRIX_CONFIGURATIONS
    return [
        (cfg["situation"], cfg["configuration"],
         ", ".join(cfg["codes"]), cfg["implication"])
        for cfg in MATRIX_CONFIGURATIONS
    ]


def _repr_codebook_rows():
    from zbd_common import CODEBOOK, VISIBILITY_SCALE
    rows = []
    for dim, items in CODEBOOK.items():
        for c, l in items:
            rows.append(("نشانهٔ بدنی", c, l))
    rows.append(("قاب‌بندی", "—", "کلوزآپ / نیم‌نما / تمام‌نما / پهن"))
    rows.append(("زاویهٔ دید", "—", "سطح / بالا / پایین / کج"))
    rows.append(("اندازهٔ نما", "—", "نزدیک / میانی / دور"))
    rows.append(("برش", "—", "چپ / راست / بالا / پایین / مرکزی"))
    rows.append(("اندازه در صفحه", "—", "خیلی بزرگ تا خیلی کوچک"))
    rows.append(("جایگاه در صفحه", "—", "۹ ناحیهٔ شبکهٔ ۳×۳"))
    rows.append(("مقیاس دیداری", "0–4", "ناچیز تا بسیار برجسته"))
    return rows


if __name__ == "__main__":
    from zbd_common import SH_OUT
    p1 = build_article_docx(1, SH_OUT / "اسکلت_مقاله۱.docx")
    p2 = build_article_docx(2, SH_OUT / "اسکلت_مقاله۲.docx")
    assert p1.exists() and p2.exists()
    print("zbd_docx خودآزمون: OK")
    print(p1)
    print(p2)
