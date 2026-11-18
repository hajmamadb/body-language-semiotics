# -*- coding: utf-8 -*-
"""
zbd_analyze.py — تحلیل و گزارش مشترک دو مقاله
===============================================
1) stats      : خوانش برگهٔ کدگذاری xlsx -> آمار تواتر + جدول متقابل (xlsx)
2) charts     : نمودار میله‌ای/میله‌های گروهی از آمار (PNG با فونت فارسی)
3) report     : ساخت اسکلت مقالۀ کامل (docx) با بخش‌ها و جدول‌ها
4) crosscheck : اعتبارسنجی کدها در برگهٔ کدگذاری (گزارش خطا)

CLI نمونه:
    python zbd_analyze.py stats --sheet "article1/output/کدگذاری.xlsx" --by-context
    python zbd_analyze.py charts --stats "article1/output/stats.json"
    python zbd_analyze.py report --article 1 --out "article1/output/مقاله.docx"
    python zbd_analyze.py crosscheck --sheet "article1/output/کدگذاری.xlsx"
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

from zbd_common import (CODEBOOK, ROOT, now_ir, read_json, safe_name,
                       write_json)


# ------------------------------------------------------------------ خوانش
def _load_sheet_rows(xlsx: Path) -> list[list]:
    """خوانش سطرهای یک برگهٔ xlsx (بدون سردسترس)."""
    from openpyxl import load_workbook
    wb = load_workbook(xlsx, data_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in r):
            continue
        rows.append(list(r))
    return rows


def _header_index(xlsx: Path, *needles: str) -> int:
    """شمارهٔ ستون (۰-محور) با عنوانی که حاوی هرکدام از needles باشد."""
    from openpyxl import load_workbook
    ws = load_workbook(xlsx).active
    headers = [str(c.value or "") for c in ws[1]]
    for i, h in enumerate(headers):
        if all(n in h for n in needles):
            return i
    raise SystemExit(f"ستونی با عنوان {needles} در {xlsx.name} یافت نشد")


# ------------------------------------------------------------------ stats
def stats(sheet: Path, group_by: str | None = None) -> dict:
    """تواتر کدهای هر مقولهٔ ماتریسی + (اختیاری) تقسیم بر موقعیت/روزنامه."""
    dim_cols = {
        "gaze": ["بردار", "نگاه"],
        "face": ["ریزآرایش", "چهره"],
        "head": ["انحراف", "سر"],
        "torso": ["راستای", "فقرات"],
        "left_hand": ["دست", "چپ"],
        "right_hand": ["دست", "راست"],
        "clothing": ["مرزبندی", "پوشش"],
    }
    rows = _load_sheet_rows(sheet)
    valid = codebook_valid()
    out: dict = {"generated_at": now_ir(), "source": str(sheet),
                 "n_events": len(rows), "dims": {}}
    # ستون گروه‌بندی
    gidx = None
    if group_by:
        from openpyxl import load_workbook
        ws = load_workbook(sheet).active
        headers = [str(c.value or "") for c in ws[1]]
        for i, h in enumerate(headers):
            if group_by in h:
                gidx = i
                break
        if gidx is None:
            print(f"  ! ستون گروه‌بندی '{group_by}' یافت نشد")
    for dim, needles in dim_cols.items():
        try:
            cidx = _header_index(sheet, *needles)
        except SystemExit:
            continue
        per_code = Counter()
        per_group: dict[str, Counter] = {}
        skipped = 0
        for r in rows:
            v = r[cidx]
            if v is None:
                continue
            v = str(v).strip()
            if v in valid:
                per_code[v] += 1
                if gidx is not None and r[gidx] is not None:
                    g = str(r[gidx]).strip()
                    per_group.setdefault(g, Counter())[v] += 1
            else:
                skipped += 1
        if per_code:
            out["dims"][dim] = {
                "counts": dict(per_code.most_common()),
                "total": sum(per_code.values()),
                "invalid_values": skipped,
                "by_group": {g: dict(c.most_common())
                             for g, c in sorted(per_group.items())} if gidx is not None else None,
            }
    out_path = sheet.with_name(f"stats_{sheet.stem}.json")
    write_json(out_path, out)
    # جدول آمار به xlsx
    xlsx_path = sheet.with_name(f"آمار_{sheet.stem}.xlsx")
    _stats_xlsx(out, xlsx_path)
    print("آمار در:", out_path)
    print("جدول آمار:", xlsx_path)
    return out


def codebook_valid() -> set[str]:
    return {c for dim in CODEBOOK.values() for c, _ in dim}


def _stats_xlsx(out: dict, path: Path) -> None:
    from zbd_sheet import SheetBuilder
    sb = SheetBuilder("آمار")
    sb.set_columns(["بُعد", "کد", "عنوان", "تواتر", "درصد", "گروه", "تواتر در گروه"],
                   [16, 8, 34, 10, 10, 22, 14])
    labels = {}
    for dim in CODEBOOK.values():
        labels.update(dim)
    for dim, d in out.get("dims", {}).items():
        tot = d["total"] or 1
        for code, n in d["counts"].items():
            if d.get("by_group"):
                for g, gc in d["by_group"].items():
                    sb.add_row(dim, code, labels.get(code, ""), n,
                               f"{n / tot * 100:.1f}%", g, gc.get(code, 0))
            else:
                sb.add_row(dim, code, labels.get(code, ""), n,
                           f"{n / tot * 100:.1f}%", "", "")
    sb.save(path)


# ------------------------------------------------------------------ charts
def charts(stats_path: Path, out_dir: Path | None = None) -> list[Path]:
    """نمودار میله‌ای تواتر هر بُعد (با برچسب فارسی)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    stats = read_json(stats_path)
    out_dir = Path(out_dir) if out_dir else stats_path.parent / "charts"
    out_dir.mkdir(parents=True, exist_ok=True)
    # فونت: Tahoma هم فارسی دارد هم لاتین (برای کدهای G1/F2) -> امن‌ترین انتخاب
    fa = None
    for fp in ("C:/Windows/Fonts/tahoma.ttf",
               "C:/Windows/Fonts/BNazanin.ttf",
               "C:/Windows/Fonts/Nazanin.ttf"):
        p = Path(fp)
        if p.exists():
            font_manager.fontManager.addfont(str(p))
            fa = font_manager.FontProperties(fname=str(p))
            break
    if fa:
        plt.rcParams["font.family"] = fa.get_name()
    plt.rcParams["axes.unicode_minus"] = False

    dim_names = {
        "gaze": "بردار نگاه",
        "face": "ریزآرایش چهره",
        "head": "انحراف محوری سر",
        "torso": "راستای ستون فقرات",
        "left_hand": "دامنهٔ دست چپ (کنش‌گر)",
        "right_hand": "وضعیت دست راست (تکیه‌گاه)",
        "clothing": "مرزبندی پوشش (Framing)",
    }
    saved = []
    for dim, d in stats.get("dims", {}).items():
        items = list(d["counts"].items())
        if not items:
            continue
        labels = [c for c, _ in items]
        vals = [n for _, n in items]
        fig, ax = plt.subplots(figsize=(8, max(3, 0.5 * len(items) + 1.5)))
        ax.barh(labels[::-1], vals[::-1])
        ax.set_title(f"{dim_names.get(dim, dim)} — تواتر نشانه‌ها", fontproperties=fa)
        for i, v in enumerate(vals[::-1]):
            ax.text(v, i, str(v), va="center", ha="left", fontsize=9)
        ax.set_xlabel("تواتر")
        p = out_dir / f"chart_{dim}.png"
        fig.tight_layout()
        fig.savefig(p, dpi=150)
        plt.close(fig)
        saved.append(p)
    # نمودار گروهی موقعیت‌ها (اگر باشد)
    for dim, d in stats.get("dims", {}).items():
        bg = d.get("by_group")
        if not bg:
            continue
        groups = list(bg)
        fig, ax = plt.subplots(figsize=(9, max(4, 0.8 * len(groups) + 2)))
        codes = sorted({c for cc in bg.values() for c in cc}, key=lambda x: (x[0], x[1]))
        import numpy as np
        y = np.arange(len(groups))
        h = 0.8 / len(codes)
        for k, code in enumerate(codes):
            vals = [bg[g].get(code, 0) for g in groups]
            ax.barh(y + k * h, vals, height=h * 0.9, label=code)
        ax.set_yticks(y + h * (len(codes) - 1) / 2)
        ax.set_yticklabels(groups)
        ax.legend(prop=fa, fontsize=8)
        ax.set_title(f"{dim_names.get(dim, dim)} در میان موقعیت‌ها", fontproperties=fa)
        p = out_dir / f"chart_{dim}_bygroup.png"
        fig.tight_layout()
        fig.savefig(p, dpi=150)
        plt.close(fig)
        saved.append(p)
    print(f"{len(saved)} نمودار در: {out_dir}")
    return saved


# ------------------------------------------------------------------ crosscheck
def crosscheck(sheet: Path) -> list[str]:
    """اعتبارسنجی مقادیر کد در ستون‌های کدگذاری؛ فهرست خطاها را برمی‌گرداند."""
    rows = _load_sheet_rows(sheet)
    valid = codebook_valid()
    from zbd_common import CODEBOOK
    dim_cols = {
        "gaze": ["بردار", "نگاه"],
        "face": ["ریزآرایش", "چهره"],
        "head": ["انحراف", "سر"],
        "torso": ["راستای", "فقرات"],
        "left_hand": ["دست", "چپ"],
        "right_hand": ["دست", "راست"],
        "clothing": ["مرزبندی", "پوشش"],
    }
    errors: list[str] = []
    for dim, needles in dim_cols.items():
        try:
            cidx = _header_index(sheet, *needles)
        except SystemExit:
            continue
        for i, r in enumerate(rows, start=2):
            v = r[cidx]
            if v is None or (isinstance(v, str) and not str(v).strip()):
                continue
            v = str(v).strip()
            if v not in valid:
                errors.append(f"سطر {i}، بُعد {dim}: مقدار نامعتبر {v!r}")
    return errors


# ------------------------------------------------------------------ report
def report(article: int, out: Path) -> Path:
    """ساخت اسکلت docx مقالۀ فارسی (RTL) با بخش‌ها و جداول آماده."""
    from zbd_docx import build_article_docx
    return build_article_docx(article, out)


# ------------------------------------------------------------------ CLI
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    ps = sub.add_parser("stats")
    ps.add_argument("--sheet", required=True, help="برگهٔ کدگذاری xlsx")
    ps.add_argument("--by-context", action="store_true",
                    help="تقسیم آمار بر موقعیت ارتباطی / روزنامه")

    pc = sub.add_parser("charts")
    pc.add_argument("--stats", required=True, help="فایل stats.json")
    pc.add_argument("--out")

    pr = sub.add_parser("report")
    pr.add_argument("--article", type=int, choices=[1, 2], required=True)
    pr.add_argument("--out", required=True)

    px = sub.add_parser("crosscheck")
    px.add_argument("--sheet", required=True)

    args = ap.parse_args(argv)

    if args.cmd == "stats":
        sheet = Path(args.sheet)
        if not sheet.is_absolute() and not sheet.exists():
            sheet = ROOT / sheet
        stats(sheet, group_by="موقعیت ارتباطی" if args.by_context else None)
        return 0
    if args.cmd == "charts":
        p = Path(args.stats)
        if not p.exists():
            p = ROOT / p
        charts(p, Path(args.out) if args.out else None)
        return 0
    if args.cmd == "report":
        out = Path(args.out)
        if not out.is_absolute():
            out = ROOT / out
        print("اسکلت مقاله:", report(args.article, out))
        return 0
    if args.cmd == "crosscheck":
        sheet = Path(args.sheet)
        if not sheet.exists():
            sheet = ROOT / sheet
        errs = crosscheck(sheet)
        if errs:
            print(f"{len(errs)} خطای اعتبارسنجی:")
            for e in errs[:40]:
                print("  ", e)
            return 2
        print("اعتبارسنجی: همهٔ کدها معتبرند ✔")
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
