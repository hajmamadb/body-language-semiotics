# -*- coding: utf-8 -*-
"""
zbd_download.py — فهرست و دانلود ویدیوهای مقالهٔ ۱
===================================================
این اسکریپت:
  ۱) فهرست ۲۵ ویدیوی پیشنهادی Gemini (از docx) را در یک xlsx ذخیره می‌کند.
  ۲) با yt-dlp ویدیوها را از لینک‌های واقعی khamenei.ir / aparat.com دانلود می‌کند.
  ۳) هر ویدیو را با نام‌گذاری معنادار (SIT-CODE-YYYYMM-FNN.mp4) ذخیره می‌کند.

نحوهٔ اجرا:
  python -I tools/zbd_download.py inventory    # فقط xlsx بسازد
  python -I tools/zbd_download.py download     # همه را دانلود کند
  python -I tools/zbd_download.py download --sit 1 3 5   # فقط موقعیت‌های ۱، ۳، ۵
  python -I tools/zbd_download.py download --id 27 31    # فقط رکوردهای خاص

مکان خروجی: article1/data/ویدیوها/
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

# ------------------------------------------------------------------ مسیرها
ROOT = Path(__file__).resolve().parent.parent
A1_VIDEO_DIR = ROOT / "article1" / "data" / "ویدیوها"
A1_OUTPUT = ROOT / "article1" / "output"
INVENTORY_XLSX = A1_OUTPUT / "فهرست_ویدیوها.xlsx"

A1_VIDEO_DIR.mkdir(parents=True, exist_ok=True)
A1_OUTPUT.mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------------
# کد موقعیت برای نام‌گذاری فایل‌ها
# ------------------------------------------------------------------
SIT_CODE = {
    1: "MFY",  # martyr_families
    2: "DPL",  # diplomatic_foreign
    3: "OFF",  # domestic_officials
    4: "MRP",  # military_parade
    5: "POT",  # poet_meeting
}
SIT_NAME = {
    1: "شهدا", 2: "دیپل", 3: "کارگزاران", 4: "رژه", 5: "شاعران"
}

# ------------------------------------------------------------------
# ۲۵ رکورد ویدیویی — استخراج‌شده از docx Gemini
# (situation, num, title, date_approx, source, url, duration, res, frame_desc)
# ------------------------------------------------------------------
RECORDS: list[dict] = [
    # ---- موقعیت ۱: خانواده شهدا ----
    dict(situation=1, num=1,
         title="دیدار خانواده شهید سپهبد حاج قاسم سلیمانی",
         date="دی ۱۳۹۹", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=49262",
         duration="کامل", res="720p/1080p",
         frame="نشسته، خانواده مقابل، تمایل تنه به جلو، نگاه متمرکز به پایین"),
    dict(situation=1, num=2,
         title="دیدار خانواده شهید محسن حججی",
         date="مهر ۱۳۹۶", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=37857",
         duration="کامل", res="720p",
         frame="نشسته، دیدار صمیمی، دست‌دادن احتمالی"),
    dict(situation=1, num=3,
         title="دیدار منزل حسین همدانی",
         date="مهر ۱۳۹۴", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=50764",
         duration="کامل", res="720p",
         frame="محیط منزل، قاب صمیمی، تماس فیزیکی"),
    dict(situation=1, num=4,
         title="دیدار منزل مسیحی (آرمن آودیسیان)",
         date="دی ۱۳۹۶", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=38627",
         duration="کامل", res="720p",
         frame="محیط منزل، تعاملات نزدیک"),
    dict(situation=1, num=5,
         title="دیدار و گفت‌وگو با خانواده شهید روح‌الله عجمیان",
         date="نامشخص", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=52611",
         duration="کامل", res="720p",
         frame="نشسته، گفت‌وگو، لمس دست"),

    # ---- موقعیت ۲: دیپلماتیک ----
    dict(situation=2, num=1,
         title="دیدار رئیس‌جمهور روسیه (پوتین)",
         date="تیر ۱۴۰۱ (یا ۱۳۹۴)", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=55940",
         duration="کامل", res="720p/1080p",
         frame="میز مذاکره، تنه زاویه‌دار، نگاه مستقیم به چشم"),
    dict(situation=2, num=2,
         title="دیدار رئیس‌جمهور سوریه (بشار اسد)",
         date="اسفند ۱۳۹۷", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=41797",
         duration="کامل", res="720p",
         frame="نشست دو‌طرفه، قاب رسمی"),
    dict(situation=2, num=3,
         title="دیدار نخست‌وزیر ژاپن (شینزو آبه)",
         date="خرداد ۱۳۹۸", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=42838",
         duration="کامل", res="720p",
         frame="نشست رسمی، دست چپ روی دسته صندلی"),
    dict(situation=2, num=4,
         title="«خط دیدار» سلطان عمان (هیثم بن طارق) و هیئت همراه",
         date="خرداد ۱۴۰۲", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=52981",
         duration="کامل", res="1080p",
         frame="قاب رسمی، تنه جبهه‌ای"),
    dict(situation=2, num=5,
         title="نماهنگ | اخراج آمریکا از منطقه (دیدار قطر - شیخ تمیم)",
         date="دی ۱۳۹۸", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=44662",
         duration="گزیده", res="720p",
         frame="نشست دو‌طرفه، قاب رسمی"),

    # ---- موقعیت ۳: کارگزاران ----
    dict(situation=3, num=1,
         title="دیدار رمضانی با مسئولان (۵۶ دقیقه)",
         date="فروردین ۱۴۰۳", source="آپارات",
         url="https://www.aparat.com/v/jvh1v6l",
         duration="۵۶ دقیقه", res="1080p/720p",
         frame="نشسته در جمع، نگاه سیال در جمع"),
    dict(situation=3, num=2,
         title="بیانات در دیدار رئیس و نمایندگان مجلس دوازدهم",
         date="تیر ۱۴۰۳", source="آپارات",
         url="https://www.aparat.com/v/ngt2ey3",
         duration="۳۸ دقیقه", res="720p",
         frame="نشسته پشت میز، اشاره با دست چپ"),
    dict(situation=3, num=3,
         title="نخستین دیدار هیئت دولت سیزدهم",
         date="شهریور ۱۴۰۰", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=48589",
         duration="کامل", res="720p",
         frame="نشسته در جمع کارگزاران"),
    dict(situation=3, num=4,
         title="بیانات در دیدار رئیس‌جمهور و اعضای هیئت دولت",
         date="تیر ۱۳۹۷", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=41169",
         duration="کامل", res="720p",
         frame="نشسته پشت میز، نگاه به جمع"),
    dict(situation=3, num=5,
         title="دیدار کارگزاران حج",
         date="خرداد ۱۴۰۱", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=52835",
         duration="کامل", res="720p",
         frame="نشسته، تعامل با کارگزاران"),

    # ---- موقعیت ۴: رژه نظامی ----
    dict(situation=4, num=1,
         title="مراسم مشترک دانش‌آموختگی دانشجویان دانشگاه‌های افسری",
         date="مهر ۱۴۰۲ (یا ۱۴۰۰)", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=48784",
         duration="کامل", res="1080p/720p",
         frame="ایستاده بر سکو، نگاه به افق، انجماد عضلات"),
    dict(situation=4, num=2,
         title="بیانات در مراسم دانش‌آموختگی دانشجویان دانشگاه افسری ارتش",
         date="آبان ۱۳۹۶", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=38005",
         duration="کامل", res="720p",
         frame="ایستاده بر سکو، چانه افراشته"),
    dict(situation=4, num=3,
         title="حضور و سخنرانی در مراسم دانش‌آموختگی دانشگاه امام حسین (ع)",
         date="مهر ۱۳۹۸", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=43776",
         duration="کامل", res="720p",
         frame="ایستاده، پوشش روحانی در کنار نظامیان"),
    dict(situation=4, num=4,
         title="حضور و سخنرانی در دانشگاه علوم دریایی نوشهر",
         date="شهریور ۱۳۹۷", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=40455",
         duration="کامل", res="720p",
         frame="ایستاده بر سکو، رژه دریایی"),
    dict(situation=4, num=5,
         title="مراسم مشترک دانش‌آموختگی (پاندمی - مهر ۱۴۰۰)",
         date="مهر ۱۴۰۰", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=54050",
         duration="کامل", res="720p",
         frame="ایستاده، محدودیت پاندمی"),

    # ---- موقعیت ۵: شاعران ----
    dict(situation=5, num=1,
         title="دیدار شاعران و اهالی فرهنگ و ادب فارسی",
         date="فروردین ۱۴۰۲", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=59690",
         duration="کامل", res="1080p/720p",
         frame="مدیوم‌کلوزآپ، دوربین ثابت، تکان‌های سر و لبخندها"),
    dict(situation=5, num=2,
         title="دیدار شاعران و اساتید ادبیات فارسی",
         date="رمضان ۱۴۰۳ (یا اسفند)", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=52429",
         duration="کامل/گزیده", res="720p",
         frame="نمای باز مایل، تقرب تنه به جلو، دریافت کتاب"),
    dict(situation=5, num=3,
         title="دیدار شاعران در شب میلاد امام حسن (ع)",
         date="فروردین ۱۴۰۱", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=55837",
         duration="کامل", res="720p",
         frame="مدیوم‌شات، تنوع واکنش‌های چهره"),
    dict(situation=5, num=4,
         title="دیدار جمعی از شاعران و اهالی فرهنگ و ادب",
         date="خرداد ۱۳۹۷", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=39751",
         duration="کامل (حجم بالا)", res="1080p/720p",
         frame="قاب صمیمی، پان کند، کینم‌های پذیرا"),
    dict(situation=5, num=5,
         title="دیدار جمعی از شاعران آئینی و مداحان خراسان",
         date="تیر ۱۳۹۵", source="Khamenei.ir",
         url="https://farsi.khamenei.ir/video-content?id=33811",
         duration="گزیده/کامل", res="720p",
         frame="نمای نزدیک‌تر، محیط حرم رضوی"),
]

# ------------------------------------------------------------------
# ساخت نام فایل معنادار
# ------------------------------------------------------------------
def make_filename(rec: dict) -> str:
    """SIT-CODE-NN.mp4 — نام ASCII بدون مشکل انکودینگ."""
    sit_code = SIT_CODE[rec["situation"]]
    fnn = f"{rec['num']:02d}"
    return f"SIT-{sit_code}-{fnn}.mp4"


# ------------------------------------------------------------------
# ساخت xlsx فهرست
# ------------------------------------------------------------------
def build_inventory():
    """فهرست xlsx را می‌سازد."""
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("خطا: openpyxl نصب نیست. pip install openpyxl")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "فهرست ویدیوها"

    headers = [
        "ردیف", "موقعیت", "کد موقعیت", "عنوان/موضوع رویداد",
        "تاریخ (تقریبی)", "منبع", "لینک", "مدت", "رزولوشن",
        "توضیح قاب", "نام فایل", "وضعیت دانلود"
    ]
    ws.append(headers)

    # سبز روشن برای هدر
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    for col_idx, _ in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # عرض ستون‌ها
    widths = [6, 10, 12, 50, 22, 14, 60, 14, 14, 50, 45, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, rec in enumerate(RECORDS, 1):
        sit_name = {
            1: "خانواده شهدا", 2: "دیپلماتیک", 3: "کارگزاران",
            4: "رژه نظامی", 5: "شاعران"
        }[rec["situation"]]
        row = [
            idx,
            sit_name,
            SIT_CODE[rec["situation"]],
            rec["title"],
            rec["date"],
            rec["source"],
            rec["url"],
            rec["duration"],
            rec["res"],
            rec["frame"],
            make_filename(rec),
            "—"  # وضعیت: قبل از دانلود
        ]
        ws.append(row)

    # فیلتر
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(INVENTORY_XLSX)
    print(f"✓ فهرست در {INVENTORY_XLSX} ذخیره شد ({len(RECORDS)} رکورد)")


# ------------------------------------------------------------------
# دانلود با yt-dlp
# ------------------------------------------------------------------
_YTDLP_KIND: str | None = None  # "cli" | "module" | None — یک‌بار تشخیص، بعدأ کش


def _detect_ytdlp() -> str:
    """تشخیص نوع دسترسی: cli یا module. در غیر این صورت None."""
    global _YTDLP_KIND
    if _YTDLP_KIND is not None:
        return _YTDLP_KIND
    # اول باینری CLI
    try:
        r = subprocess.run(["yt-dlp", "--version"],
                           capture_output=True, text=True, timeout=15)
        if r.returncode == 0:
            _YTDLP_KIND = "cli"
            return _YTDLP_KIND
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass
    # بعد ماژول Python
    try:
        r = subprocess.run([sys.executable, "-m", "yt_dlp", "--version"],
                           capture_output=True, text=True, timeout=15)
        if r.returncode == 0:
            _YTDLP_KIND = "module"
            return _YTDLP_KIND
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass
    return "none"


def check_ytdlp() -> bool:
    """yt-dlp موجود است؟"""
    kind = _detect_ytdlp()
    if kind == "cli":
        print("yt-dlp: CLI")
        return True
    if kind == "module":
        print("yt-dlp: module (python -m yt_dlp)")
        return True
    return False


def ytdlp_cmd() -> list[str]:
    """دستور yt-dlp (CLI یا module)."""
    kind = _detect_ytdlp()
    if kind == "cli":
        return ["yt-dlp"]
    return [sys.executable, "-m", "yt_dlp"]


def download_one(rec: dict, dest_dir: Path) -> tuple[bool, str]:
    """یک ویدیو دانلود می‌کند."""
    url = rec["url"]
    filename = make_filename(rec)
    dest_path = dest_dir / filename

    # اگر قبلاً دانلود شده
    if dest_path.exists():
        size_mb = dest_path.stat().st_size / (1024 * 1024)
        return True, f"قبلاً دانلود شده ({size_mb:.1f} MB)"

    cmd = ytdlp_cmd() + [
        "-f", "bestvideo[height<=1080]+bestaudio/best[height<=1080]/best",
        "--merge-output-format", "mp4",
        "-o", str(dest_path),
        "--no-playlist",
        "--write-info-json",
        "--no-warnings",
        "--socket-timeout", "30",
        "--retries", "3",
        url
    ]

    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if r.returncode == 0:
            if dest_path.exists():
                size_mb = dest_path.stat().st_size / (1024 * 1024)
                return True, f"OK ({size_mb:.1f} MB)"
            # شاید نام فایل متفاوت بوده
            mp4s = list(dest_dir.glob("*.mp4"))
            if mp4s:
                largest = max(mp4s, key=lambda p: p.stat().st_size)
                if largest != dest_path:
                    largest.rename(dest_path)
                size_mb = dest_path.stat().st_size / (1024 * 1024)
                return True, f"OK ({size_mb:.1f} MB) [rename]"
            return False, "فایل پیدا نشد"
        else:
            err = r.stderr.strip()[:200] if r.stderr else "unknown"
            return False, f"خطا: {err}"
    except subprocess.TimeoutExpired:
        return False, "تایم‌اوت (۱۰ دقیقه)"
    except Exception as e:
        return False, f"خطای ناگهانی: {e}"


def download_all(situations: set[int] | None = None,
                 record_ids: set[int] | None = None):
    """همه ویدیوها را دانلود می‌کند."""
    if not check_ytdlp():
        print("خطا: yt-dlp نصب نیست. ابتدا: pip install yt-dlp")
        return

    total = len(RECORDS)
    done = 0
    failed = 0
    skipped = 0

    for idx, rec in enumerate(RECORDS, 1):
        # فیلتر موقعیت
        if situations and rec["situation"] not in situations:
            continue
        # فیلتر id
        if record_ids and idx not in record_ids:
            continue

        sit_name = SIT_NAME[rec["situation"]]
        label = f"[{idx}/{total}] موقعیت {rec['situation']} ({sit_name}) #{rec['num']} — {rec['title']}"
        print(f"\n{'='*60}")
        print(label)
        print(f"  URL: {rec['url']}")

        ok, msg = download_one(rec, A1_VIDEO_DIR)
        if ok:
            done += 1
            print(f"  ✓ {msg}")
        else:
            failed += 1
            print(f"  ✗ {msg}")

    print(f"\n{'='*60}")
    print(f"نتیجه: {done} موفق، {failed} ناموفق")
    print(f"مکان ذخیره: {A1_VIDEO_DIR}")


# ------------------------------------------------------------------
# گزارش وضعیت فعلی فایل‌های دانلودشده
# ------------------------------------------------------------------
def report():
    """فایل‌های موجود در article1/data/ویدیوها را لیست می‌کند."""
    files = sorted(A1_VIDEO_DIR.glob("*.mp4"))
    print(f"\nفایل‌های موجود در {A1_VIDEO_DIR}:")
    if not files:
        print("  (خالی)")
    total_mb = 0
    for f in files:
        size_mb = f.stat().st_size / (1024 * 1024)
        total_mb += size_mb
        print(f"  {f.name}  ({size_mb:.1f} MB)")
    print(f"  ─────────────────────")
    print(f"  {len(files)} فایل، مجموع {total_mb:.1f} MB")


# ------------------------------------------------------------------
# main
# ------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]

    if cmd == "inventory":
        build_inventory()

    elif cmd == "download":
        situations = None
        record_ids = None
        i = 2
        while i < len(sys.argv):
            arg = sys.argv[i]
            if arg == "--sit":
                situations = set()
                i += 1
                while i < len(sys.argv) and sys.argv[i].isdigit():
                    situations.add(int(sys.argv[i]))
                    i += 1
            elif arg == "--id":
                record_ids = set()
                i += 1
                while i < len(sys.argv) and sys.argv[i].isdigit():
                    record_ids.add(int(sys.argv[i]))
                    i += 1
            else:
                i += 1
        download_all(situations, record_ids)

    elif cmd == "report":
        report()

    elif cmd == "list":
        for idx, rec in enumerate(RECORDS, 1):
            sit_name = SIT_NAME[rec["situation"]]
            fn = make_filename(rec)
            exists = "✓" if (A1_VIDEO_DIR / fn).exists() else "✗"
            print(f"{exists} [{idx:2d}] {sit_name} #{rec['num']}  {fn}")

    else:
        print(f"دستور نامعتبر: {cmd}")
        print("دستورها: inventory | download | report | list")
        print("  download --sit 1 3 5")
        print("  download --id 1 5 10")


if __name__ == "__main__":
    main()
