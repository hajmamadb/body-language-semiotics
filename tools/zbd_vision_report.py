# -*- coding: utf-8 -*-
"""
zbd_vision_report.py — تجزیهٔ خروجی YOLO برای بازنگری کتاب‌کد
==============================================================
از vision_photos.json / vision_videos.json می‌خواند و هر رکورد را به
دسته‌های «قابل‌قضاوت از داده» تبدیل می‌کند تا کدهای کتاب‌کد دقیق‌تر
و متنوع‌تر شوند. فقط ویژگی‌های قابل‌اندازه‌گیری می‌سازد:

  - shot_size  : اندازهٔ نما (کلوزآپ / نیم‌نما / تمام‌نما / پهن) از head_top_rel
  - layout     : چیدمان (تنها / جفت / گروه / پرجمعیت / صحنه)
  - vert_zone  : جایگاه عمودی مرکزِ شخص اصلی (بالا/میانی/پایین)
  - head_ratio : برآورد نسبت ارتفاع سر به بدن از aspect شخص
  - aspect_cls : کشیده/نرم/پهن بودن قاب شخص اصلی

خروجی: vision_report.json + جدول خلاصه در ترمینال (فارسی).
CLI نمونه:
    python zbd_vision_report.py --photos
    python zbd_vision_report.py --videos
    python zbd_vision_report.py --both
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

from zbd_common import A1_OUT, now_ir, read_json, write_json


# ------------------------------------------------------------------ دسته‌بندی
def classify_shot(head_top_rel: float | None) -> str:
    """اندازۀ نما بر اساس فاصلهٔ بالای سر تا بالای قاب (کمتر = کلوزآپ‌تر)."""
    if head_top_rel is None:
        return "نامشخص"
    if head_top_rel <= 0.08:
        return "کلوزآپ / سر-شانه"
    if head_top_rel <= 0.20:
        return "نیم‌نمای بالا (سینه به بالا)"
    if head_top_rel <= 0.38:
        return "نیم‌نمای پایین / زانو به بالا"
    return "تمام‌نما / پهن (کامل یا صحنه)"


def classify_layout(n_person: int) -> str:
    if n_person <= 1:
        return "تنها"
    if n_person == 2:
        return "جفت"
    if n_person <= 4:
        return "گروه کوچک (۳-۴)"
    if n_person <= 8:
        return "گروه متوسط (۵-۸)"
    return "پرجمعیت / توده (۹+)"


def classify_vert(cy_rel: float | None) -> str:
    if cy_rel is None:
        return "نامشخص"
    if cy_rel < 0.40:
        return "بالای قاب"
    if cy_rel <= 0.60:
        return "میانهٔ قاب"
    return "پایین قاب"


def classify_aspect(a: float | None) -> str:
    """نسبت عرض/ارتفاعِ قاب شخص اصلی؛ نشاเนٔ ایستاده در مقابل نشسته/پهن."""
    if a is None:
        return "نامشخص"
    if a < 0.45:
        return "کشیده (ایستاده/پوشیده)"
    if a <= 0.80:
        return "نرم (معمول ایستاده)"
    if a <= 1.20:
        return "مربعی (نشسته/کوتاه)"
    return "پهن (چندجسمی/صحنه‌ای)"


def head_ratio_from_aspect(a: float | None) -> str:
    """برآورد نسبت سر به بدن از aspect: هرچه کشیده‌تر، سر نسبتا کوچک‌تر دیده می‌شود."""
    if a is None:
        return "—"
    if a < 0.45:
        return "سر کوچک/دور (≈۱:۷+)"
    if a <= 0.80:
        return "سر متوسط (≈۱:۶)"
    if a <= 1.20:
        return "سر بزرگ/نزدیک (≈۱:۴)"
    return "چندجسمی/نامشخص"


def _record(rec: dict) -> dict:
    p = rec.get("primary_person") or {}
    cnt = rec.get("counts", {})
    return {
        "file": rec.get("file"),
        "video": rec.get("video"),
        "time_hms": rec.get("time_hms"),
        "n_person": cnt.get("person", 0),
        "shot_size": classify_shot(rec.get("head_top_rel")),
        "layout": classify_layout(cnt.get("person", 0)),
        "vert_zone": classify_vert(p.get("cy_rel")),
        "aspect_cls": classify_aspect(p.get("aspect")),
        "head_ratio": head_ratio_from_aspect(p.get("aspect")),
        "area_rel": p.get("area_rel"),
        "cx_rel": p.get("cx_rel"),
    }


# ------------------------------------------------------------------ گزارش
def _summarize(items: list[dict]) -> dict:
    def cnt(key):
        return dict(Counter(i[key] for i in items).most_common())
    return {
        "n": len(items),
        "shot_size": cnt("shot_size"),
        "layout": cnt("layout"),
        "vert_zone": cnt("vert_zone"),
        "aspect_cls": cnt("aspect_cls"),
        "head_ratio": cnt("head_ratio"),
    }


def build(photos: Path | None, videos: Path | None) -> Path:
    sections: dict = {"generated_at": now_ir()}
    all_recs: list[dict] = []

    if photos and photos.exists():
        data = read_json(photos)
        recs = [_record(r) for r in data["records"] if "error" not in r]
        sections["photos"] = {"summary": _summarize(recs), "records": recs}
        all_recs += recs

    if videos and videos.exists():
        data = read_json(videos)
        recs = [_record(r) for r in data["records"] if "error" not in r]
        sections["videos"] = {"summary": _summarize(recs), "records": recs}
        all_recs += recs

    if all_recs:
        sections["combined"] = _summarize(all_recs)

    out = A1_OUT / "vision_report.json"
    write_json(out, sections)
    _print_report(sections)
    print("خروجی:", out)
    return out


def _print_report(sections: dict) -> None:
    def show(name: str, summ: dict | None):
        if not summ:
            return
        print(f"\n{'=' * 60}\n{name}  (n={summ['n']})\n{'=' * 60}")
        for key, fa in (("shot_size", "اندازۀ نما"),
                        ("layout", "چیدمان افراد"),
                        ("vert_zone", "جایگاه عمودی"),
                        ("aspect_cls", "شیب قاب شخص اصلی"),
                        ("head_ratio", "برآورد سر/بدن")):
            print(f"  • {fa}:")
            for k, v in summ.get(key, {}).items():
                print(f"       {k:28s} {v}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--photos", action="store_true")
    ap.add_argument("--videos", action="store_true")
    ap.add_argument("--both", action="store_true")
    args = ap.parse_args(argv)

    photos = A1_OUT / "vision_photos.json" if (args.photos or args.both or not (args.photos or args.videos)) else None
    videos = A1_OUT / "vision_videos.json" if (args.videos or args.both or not (args.photos or args.videos)) else None
    build(photos, videos)
    return 0


if __name__ == "__main__":
    sys.exit(main())
