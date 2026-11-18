# -*- coding: utf-8 -*-
"""
zbd_sheet.py — کدر مشترک Excel/CSV راست‌به‌چپ برای جدول‌های کدگذاری
=====================================================================
از openpyxl برای ساخت برگه‌های xlsx با جهت RTL، قلم فارسی،
سردسترس قفل‌شده و فهرست‌های داده (Data Validation) استفاده می‌کند.
CSV خالص RTL نیست، بنابراین خروجی اصلی xlsx است و CSV نسخهٔ ساده است.
"""
from __future__ import annotations

from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from zbd_common import (CODEBOOK, CODING_GUIDELINES, MATRIX_CONFIGURATIONS,
                        VISIBILITY_SCALE, now_ir, write_json)

FA_FONT = "B Nazanin"          # قلم فارسی رایج در محیط‌های آکادمیک
FA_FALLBACK = "Tahoma"
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name=FA_FALLBACK, bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name=FA_FALLBACK, size=11)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class SheetBuilder:
    """سازندهٔ یک برگهٔ کدگذاری RTL."""

    def __init__(self, title: str):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = title[:31] or "Sheet"
        self.columns: list[str] = []
        self.rows: list[list] = []
        self._col_widths: dict[int, float] = {}
        self._validations: dict[int, list[str]] = {}   # idx ستون -> مقادیر مجاز
        self._wrap_cols: set[int] = set()

    # -- ساختار ----------------------------------------------------------
    def set_columns(self, headers: Sequence[str],
                    widths: Sequence[float] | None = None) -> "SheetBuilder":
        self.columns = list(headers)
        if widths:
            for i, w in enumerate(widths, start=1):
                self._col_widths[i] = float(w)
        return self

    def add_row(self, *values) -> "SheetBuilder":
        if self.columns and len(values) != len(self.columns):
            raise ValueError(
                f"تعداد مقادیر ({len(values)}) با تعداد ستون‌ها "
                f"({len(self.columns)}) برابر نیست."
            )
        self.rows.append(list(values))
        return self

    def add_validation(self, col_index: int, options: Sequence[str]) -> "SheetBuilder":
        """فهرست دادهٔ کشویی برای یک ستون (۱-محور)."""
        opts = [str(o) for o in options]
        joined = ",".join(opts)
        if len(joined) > 255:
            joined = joined[:255]
        self._validations[col_index] = opts
        return self

    def wrap(self, *col_indexes: int) -> "SheetBuilder":
        for i in col_indexes:
            self._wrap_cols.add(i)
        return self

    # -- خروجی -----------------------------------------------------------
    def save(self, path: Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        ws = self.ws
        ws.sheet_view.rightToLeft = True
        n_cols = max(len(self.columns), max((len(r) for r in self.rows), default=0))

        # سردسترس
        for c, name in enumerate(self.columns, start=1):
            cell = ws.cell(row=1, column=c, value=name)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER

        # بدنه
        for r, row in enumerate(self.rows, start=2):
            for c in range(n_cols):
                val = row[c] if c < len(row) else None
                cell = ws.cell(row=r, column=c + 1, value=val)
                cell.font = BODY_FONT
                cell.border = BORDER
                cell.alignment = Alignment(
                    horizontal="right", vertical="center",
                    wrap_text=(c + 1) in self._wrap_cols,
                )

        # عرض ستون‌ها
        for c in range(1, n_cols + 1):
            letter = get_column_letter(c)
            ws.column_dimensions[letter].width = self._col_widths.get(c, 16)

        # اعتبارسنجی داده
        if self.rows:
            last = len(self.rows) + 1
            for idx, opts in self._validations.items():
                letter = get_column_letter(idx)
                dv = DataValidation(
                    type="list", formula1=f'"{",".join(opts)}"', allow_blank=True
                )
                dv.error = "مقدار خارج از فهرست است"
                dv.errorTitle = "خطای ورودی"
                ws.add_data_validation(dv)
                dv.add(f"{letter}2:{letter}{last}")

        ws.freeze_panes = "A2"
        self.wb.save(path)
        return path


# ------------------------------------------------------------------
# پیش‌ساز جدول کدگذاری مقالهٔ اول
# ------------------------------------------------------------------
def article1_coding_sheet(context_keys=None) -> SheetBuilder:
    """برگهٔ کدگذاری کالبد در موقعیت (مقالهٔ اول) — جدول مشاهده‌محور ده‌عاملی."""
    from zbd_common import CONTEXTS
    sb = SheetBuilder("کدگذاری کالبد در موقعیت")
    headers = [
        # شناسهٔ واحد تحلیل (فرمت: SIT-DIP-202207-F03)
        "شناسهٔ واحد تحلیل",
        "نوع موقعیت",
        "عنوان سند / ویدیو",
        "ثانیه‌شمار شروع برش",
        "ثانیه‌شمار پایان برش",
        # ۷ مقولهٔ ماتریسی
        "بردار نگاه (G)",
        "ریزآرایش چهره (F)",
        "انحراف محوری سر (H)",
        "راستای ستون فقرات (T)",
        "دامنهٔ دست چپ — کنش‌گر (L)",
        "وضعیت دست راست — تکیه‌گاه (R)",
        "مرزبندی پوشش (C)",
        # تحلیل
        "پیکربندی / هم‌نشینی کدها",
        "محور دلالت در بافت",
        "توضیح دلالت",
        "محدودیت تصویر (زاویه/کلوزآپ/تدوین)",
        "یادداشت پژوهشی",
    ]
    widths = [16, 18, 20, 12, 12, 14, 14, 14, 14, 16, 14, 14, 24, 18, 30, 18, 24]
    sb.set_columns(headers, widths)
    # موقعیت‌های پنج‌گانه
    ctx_opts = [CONTEXTS[k] for k in (context_keys or list(CONTEXTS))]
    sb.add_validation(2, ctx_opts)
    # هفت مقولهٔ ماتریسی
    for dim, col in (("gaze", 6), ("face", 7), ("head", 8),
                     ("torso", 9), ("left_hand", 10), ("right_hand", 11),
                     ("clothing", 12)):
        sb.add_validation(col, [c for c, _ in CODEBOOK[dim]])
    axes = ["اعتبار و رسمیت موقعیت", "ارتباط صمیمی/رسمی با مخاطب",
            "تأکید بر پیام", "همدلی و توجه به مخاطب خاص",
            "مدیریت فضا و زمان موقعیت", "نقش‌آفرینی موقعیتی", "سایر"]
    sb.add_validation(14, axes)
    sb.wrap(3, 13, 14, 15, 16, 17)
    return sb


def article2_coding_sheet() -> SheetBuilder:
    """برگهٔ کدگذاری تصویر صفحهٔ اول (مقالهٔ دوم)."""
    sb = SheetBuilder("کدگذاری تصویر صفحهٔ اول")
    headers = [
        "شناسهٔ تصویر", "روزنامه", "تاریخ/شماره", "رویداد مشترک",
        "لحظهٔ ثبت (حالت بدنی غالب)", "نگاه (کد)", "حالت چهره (کد)",
        "حرکت دست/حالت دست (کد)", "وضعیت بدن (کد)", "قاب‌بندی",
        "زاویهٔ دید", "اندازهٔ نما (نزدیک/میانی/دور)", "برش (کراپ)",
        "اندازهٔ عکس در صفحه", "جایگاه در صفحه", "نسبت با تیتر",
        "نسبت با شرح تصویر", "آرایهٔ مجاور", "محور دلالت بازنمایی",
        "توضیح دلالت", "یادداشت پژوهشی",
    ]
    widths = [10, 16, 12, 20, 22, 10, 12, 14, 12, 14, 14, 16, 14, 14, 14,
              16, 16, 16, 18, 30, 24]
    sb.set_columns(headers, widths)
    frame_opts = ["کلوزآپ/نزدیک", "نیم‌نمای", "تمام‌نمای/میانی", "پهن/دور", "پشت‌نما/از پشت", "ناشناخته"]
    sb.add_validation(10, frame_opts)
    angle_opts = ["سطح (چشمی)", "بالا (از پایین)", "پایین (از بالا)", "کج/مورب", "ناشناخته"]
    sb.add_validation(11, angle_opts)
    shot_opts = ["نزدیک", "میانی", "دور", "ناشناخته"]
    sb.add_validation(12, shot_opts)
    crop_opts = ["برش چپ", "برش راست", "برش بالا", "برش پایین", "برش مرکزی", "بدون برش محسوس", "ناشناخته"]
    sb.add_validation(13, crop_opts)
    size_opts = ["خیلی بزرگ (تمام ستون)", "بزرگ (نصف صفحه+)", "متوسط", "کوچک", "خیلی کوچک (آیکون)", "ناشناخته"]
    sb.add_validation(14, size_opts)
    pos_opts = ["بالا-راست", "بالا-وسط", "بالا-چپ", "میانه-راست", "میانه-وسط",
                "میانه-چپ", "پایین-راست", "پایین-وسط", "پایین-چپ", "ناشناخته"]
    sb.add_validation(15, pos_opts)
    axes = ["اعتبار و رسمیت", "صمیمیت/رسمیت", "تأکید بر پیام", "همدلی",
            "رهبری/نفوذ", "همراهی/یکدست بودن", "سایر"]
    sb.add_validation(19, axes)
    sb.wrap(4, 5, 19, 20, 21)
    return sb


def write_codebook_xlsx(path: Path) -> Path:
    """خروجی کتاب‌کد کامل به xlsx (مراجعهٔ پژوهشگر هنگام کدگذاری)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "کتاب‌کد"
    ws.sheet_view.rightToLeft = True
    headers = ["بُعد", "کد", "عنوان", "توضیح/مثال مشاهده‌ای"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    dim_titles = {
        "gaze": "بردار نگاه (Gaze Vector)",
        "face": "ریزآرایش چهره (Facial Micro-expression)",
        "head": "انحراف محوری سر (Head Orientation)",
        "torso": "راستای ستون فقرات (Torso Orientation)",
        "left_hand": "دامنهٔ فضایی دست چپ — دست کنش‌گر (Agent Hand)",
        "right_hand": "وضعیت دست راست — دست تکیه‌گاه (Support Hand)",
        "clothing": "مرزبندی پوشش (Framing)",
    }
    r = 2
    for dim, title in dim_titles.items():
        for code, label in CODEBOOK[dim]:
            vals = [title, code, label, ""]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = BODY_FONT
                cell.border = BORDER
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            r += 1
    # راهنمای کدگذاری (قواعد روش‌شناختی)
    ws.cell(row=r + 1, column=1,
            value="راهنمای کدگذاری (قواعد روش‌شناختی)").font = Font(bold=True)
    r += 2
    for title, body in CODING_GUIDELINES:
        c = ws.cell(row=r, column=1, value=title)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        c2 = ws.cell(row=r, column=3, value=body)
        c2.font = BODY_FONT
        c2.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        r += 1
    # ماتریس هم‌نشینی نشانه‌ها
    ws.cell(row=r + 1, column=1,
            value="ماتریس هم‌نشینی نشانه‌ها — پیکربندی‌های الگویی").font = Font(bold=True)
    r += 2
    for cfg in MATRIX_CONFIGURATIONS:
        ws.cell(row=r, column=1, value=cfg["situation"]).font = Font(bold=True)
        c = ws.cell(row=r, column=3, value=cfg["configuration"])
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        r += 1
        c = ws.cell(row=r, column=2, value=", ".join(cfg["codes"]))
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        c = ws.cell(row=r, column=4, value=cfg["implication"])
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        r += 1
        r += 1  # فاصلهٔ بصری میان پیکربندی‌ها
    # مقیاس دیداری
    ws.cell(row=r + 1, column=1, value="مقیاس دیداری/برجستگی").font = Font(bold=True)
    r += 2
    for code, label in VISIBILITY_SCALE:
        ws.cell(row=r, column=2, value=code).font = BODY_FONT
        ws.cell(row=r, column=3, value=label).font = BODY_FONT
        r += 1
    for col, w in zip("ABCD", (26, 24, 46, 42)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    from zbd_common import SH_OUT
    # خودآزمون
    p1 = SH_OUT / "test_a1.xlsx"
    article1_coding_sheet().save(p1)
    p2 = SH_OUT / "test_a2.xlsx"
    article2_coding_sheet().save(p2)
    p3 = SH_OUT / "کتاب‌کد.xlsx"
    write_codebook_xlsx(p3)
    assert p1.exists() and p2.exists() and p3.exists()
    print("zbd_sheet خودآزمون: OK")
    print("کتاب‌کد در:", p3)
