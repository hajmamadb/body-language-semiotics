# -*- coding: utf-8 -*-
"""
zbd_batch_dl.py — دانلود انبوه ویدیوهای تحقیق از CDN کhamenei.ir
==================================================================
URLها از HTML صفحه‌ها در مرورگر (IAB) استخراج شده‌اند (CDN محافظت ندارد).
برای هر ویدیو:
  1) etag/Content-Range از CDN خوانده می‌شود (سایز دقیق)
  2) اگر فایل مقصد با همان سایز وجود دارد -> skip
  3) دانلود موازی (zbd_parallel_dl) — اگر چند فایل جدا بود، با ffmpeg concat
  4) etag مقصد با CDN چک می‌شود -> verify

کاربرد:
    python -I zbd_batch_dl.py [--only 1 2 6] [--dry-run]
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zbd_parallel_dl import download_parallel, get_file_size, UA

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "article1", "data", "ویدیوها")

FFMPEG = r"C:\Users\MamaDDelta\AppData\Roaming\Python\Python315\site-packages\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v9.0.2.exe"

import requests

# ────────────────────────────────────────────────────────────────────
# URLهای CDN — از HTML صفحات در مرورگر استخراج شده (۲۰۲۶-۱۱-۰۹)
# هر entry: (situation, num, page_id, title, [guids در ترتیب پخش])
# ────────────────────────────────────────────────────────────────────
BASE = "https://idc0-cdn5.khamenei.ir/ndata/news/{id}/{guid}.mp4"

RECORDS = [
    # ── موقعیت 1: خانواده شهدا (MFY) ──
    dict(sit=1, num=1, page_id=49262, title="دیدار خانواده شهید سلیمانی",
         guids=["14001011_60366"], note="فیلم کامل (526MB)"),
    dict(sit=1, num=2, page_id=37857, title="دیدار خانواده شهید محسن حججی",
         guids=["13960711_16323", "13960711_16550", "13960711_16549",
                "13960711_16548", "13960711_16547", "13960711_16546"],
         note="6 بخش جدا — نیاز به concat"),
    dict(sit=1, num=3, page_id=50764, title="مستند کوتاه | مزد اخلاص (شهید همدانی)",
         guids=["14010520_61166"], note="فیلم کامل (209MB)"),
    dict(sit=1, num=4, page_id=38627, title="آن شب شیرین (شهید آرمن آودیسیان)",
         guids=["13961013_17663"], note="تک‌فایل"),
    dict(sit=1, num=5, page_id=52611, title="دیدار خانواده شهید عجمیان (حاشیه دیدار سفرای)",
         guids=["14020202_62281"], note="تک‌فایل"),
    # ── موقعیت 2: دیپلماسی و مسئولان خارجی (DPL) ──
    dict(sit=2, num=1, page_id=55940, title="بیانات روز جهانی قدس (مسئولان نظام)",
         guids=["14030115_64079"], note="entry انتخابی از 35 بخش"),
    dict(sit=2, num=2, page_id=41797, title="دیدار رئیس‌جمهور سوریه",
         guids=["13971206_24838"], note="تک‌فایل"),
    dict(sit=2, num=3, page_id=42838, title="دیدار نخست‌وزیر ژاپن",
         guids=["13980323_32080"], note="تک‌فایل (56MB)"),
    dict(sit=2, num=4, page_id=52981, title="«خط دیدار» دیدار سلطان عمان",
         guids=["14020308_62468"], note="تک‌فایل (21MB)"),
    dict(sit=2, num=5, page_id=44662, title="نماهنگ | اخراج آمریکا از منطقه",
         guids=["13981023_44710"], note="تک‌فایل (53MB)"),
    # ── موقعیت 3: دیدار مسئولان داخلی (OFF) ──
    dict(sit=3, num=1, page_id=48589, title="نخستین دیدار هیئت دولت سیزدهم",
         guids=["14000606_58921"], note="فیلم کامل (1.64GB)"),
    dict(sit=3, num=2, page_id=41169, title="بیانات در دیدار رئیس‌جمهور و هیئت دولت",
         guids=["13970424_22734", "13970424_24857", "13970424_24858",
                "13970424_24859", "13970424_24860", "13970424_24861",
                "13970424_24862"], note="7 بخش (761MB) — ترتیب از player"),
    dict(sit=3, num=3, page_id=52835, title="دیدار کارگزاران حج",
         guids=["14020227_62403"], note="فیلم کامل (651MB)"),
    # ── موقعیت 4: مراسم نظامی (MRP) ──
    dict(sit=4, num=1, page_id=48784, title="مراسم دانش‌آموختگی دانشجویان افسری (1400)",
         guids=["14000711_59975"], note="فیلم کامل"),
    dict(sit=4, num=2, page_id=38005, title="مراسم دانش‌آموختگی دانشجویان افسری ارتش (1396)",
         guids=["13960803_16377", "13960803_16378", "13960803_16379",
                "13960803_16380", "13960803_16381", "13960803_16382",
                "13960803_16383", "13960803_16384", "13960803_16385",
                "13960803_16386", "13960803_16387", "13960803_16388"],
         note="12 بخش — ترتیب از پلی‌لیست صفحه"),
    dict(sit=4, num=3, page_id=43776, title="مراسم دانش‌آموختگی دانشگاه امام حسین (ع) (1398)",
         guids=["13980721_42412", "13980721_44608", "13980721_44607",
                "13980721_44606", "13980721_44605", "13980721_44604",
                "13980721_44603", "13980721_44602", "13980721_44601",
                "13980721_44600", "13980721_42414"],
         note="11 بخش — ترتیب از پلی‌لیست صفحه"),
    dict(sit=4, num=4, page_id=40455, title="دانشگاه علوم دریایی امام خمینی نوشهر (1397)",
         guids=["13970618_20419", "13970618_20439", "13970618_20434",
                "13970618_20417", "13970618_20416", "13970618_20415",
                "13970618_20414", "13970618_20418"],
         note="8 بخش — ترتیب از پلی‌لیست صفحه"),
    dict(sit=4, num=5, page_id=54050, title="مراسم دانش‌آموختگی دانشجویان افسری (1402)",
         guids=["14020718_63000"], note="فیلم کامل (798MB)"),
    # ── موقعیت 5: دیدار شاعران (POT) ──
    dict(sit=5, num=1, page_id=59690, title="گزیده بیانات دیدار شاعران 1403",
         guids=["14031225_66773"], note="گزیده (4MB)"),
    dict(sit=5, num=2, page_id=52429, title="دیدار شاعران و اساتید ادبیات (1402)",
         guids=["14020116_62200"], note="فیلم کامل (618MB)"),
    dict(sit=5, num=3, page_id=55837, title="دیدار شاعران (1403)",
         guids=["14030106_64057"], note="فیلم کامل"),
    dict(sit=5, num=4, page_id=39751, title="دیدار جمعی از شاعران و اهالی فرهنگ و ادب (1397)",
         guids=["13970309_18132", "13970309_20375", "13970309_20374",
                "13970309_20373", "13970309_20372", "13970309_20371",
                "13970309_20370", "13970309_18148", "13970309_18147",
                "13970309_18146", "13970309_18145", "13970309_18144",
                "13970309_18143", "13970309_18142", "13970309_18141",
                "13970309_18140", "13970309_18139", "13970309_18138",
                "13970309_18137", "13970309_18136", "13970309_18135",
                "13970309_18134", "13970309_18133", "13970309_18130",
                "13970309_18129", "13970309_18128", "13970309_18127",
                "13970309_18126", "13970309_18125", "13970309_18124",
                "13970309_18123", "13970309_18121", "13970309_18120",
                "13970309_18119", "13970309_18117", "13970309_18116",
                "13970309_18115", "13970309_18114", "13970309_18113"],
         note="38 بخش — ترتیب از پلی‌لیست صفحه (ویدیو بلند، چند GB)"),
    dict(sit=5, num=5, page_id=33811, title="دیدار جمعی از شاعران آئینی و مداحان خراسان (1395)",
         guids=["13950424_7953", "13950424_12936", "13950424_12935",
                "13950424_12934", "13950424_12933", "13950424_8971",
                "13950424_8970", "13950424_8969", "13950424_8968",
                "13950424_8967", "13950424_8966", "13950424_8965",
                "13950424_8964", "13950424_8963", "13950424_8962",
                "13950424_8961", "13950424_8960", "13950424_8959",
                "13950424_8958", "13950424_8957", "13950424_8956",
                "13950424_8954"],
         note="22 بخش — ترتیب از پلی‌لیست صفحه (ویدیو بلند)"),
]

SIT_CODE = {1: "MFY", 2: "DPL", 3: "OFF", 4: "MRP", 5: "POT"}


def make_dest(rec) -> str:
    code = SIT_CODE[rec["sit"]]
    return os.path.join(OUT_DIR, f"SIT-{code}-{rec['num']:02d}.mp4")


def cdn_url(page_id: int, guid: str) -> str:
    return BASE.format(id=page_id, guid=guid)


def head_info(url: str) -> tuple[int, str] | None:
    """سایز کل + etag از CDN."""
    try:
        r = requests.get(url, headers={"User-Agent": UA, "Range": "bytes=0-0"},
                         timeout=20, stream=True)
        cr = r.headers.get("Content-Range", "")
        etag = r.headers.get("ETag", "")
        r.close()
        if "/" in cr:
            return int(cr.split("/")[-1]), etag
    except Exception:
        return None
    return None


def etag_matches(local_etag_hex: str, remote_etag: str) -> bool:
    """etag فرمت "hex-md5hex" دارد؛ قسمت hex با سایز فایل هم‌خوان است."""
    if not remote_etag:
        return False
    remote_hex = remote_etag.strip('"').split("-")[0].lower()
    return remote_hex == local_etag_hex.lower()


def file_md5_hex_size(path: str) -> str:
    """سایز فایل به هگزادسیمال (مطابق قسمت اول etag سرور)."""
    return format(os.path.getsize(path), "x")


def ffmpeg_concat(parts: list[str], dest: str) -> None:
    """چسباندن چند MP4 با stream copy."""
    listfile = dest + ".concat.txt"
    with open(listfile, "w", encoding="utf-8") as f:
        for p in parts:
            f.write(f"file '{p}'\n")
    cmd = [FFMPEG, "-y", "-f", "concat", "-safe", "0", "-i", listfile,
           "-c", "copy", "-movflags", "+faststart", dest]
    subprocess.run(cmd, check=True, capture_output=True)
    os.remove(listfile)


def process_one(rec: dict, dry_run: bool) -> dict:
    dest = make_dest(rec)
    result = dict(page_id=rec["page_id"], sit=rec["sit"], num=rec["num"],
                  dest=dest)
    if not rec["guids"]:
        result["status"] = "NO_URL"
        return result
    urls = [cdn_url(rec["page_id"], g) for g in rec["guids"]]

    if len(urls) == 1:
        info = head_info(urls[0]) if not dry_run else None
        if dry_run:
            result.update(status="DRY", url=urls[0])
            return result
        if not info:
            result.update(status="HEAD_FAIL", url=urls[0])
            return result
        size, etag = info
        result.update(size_mb=round(size / 1e6, 1), etag=etag)
        if os.path.exists(dest) and os.path.getsize(dest) == size:
            result["status"] = "OK_EXISTS"
            return result
        print(f"  دانلود {size/1e6:.0f}MB از {urls[0]}")
        download_parallel(urls[0], dest, workers=10, block_mb=8)
        # verify
        if os.path.exists(dest) and os.path.getsize(dest) == size:
            result["status"] = "OK"
        else:
            result["status"] = "INCOMPLETE"
        return result

    # ── چند بخش -> جداگانه دانلود و بعد concat ──
    part_files = []
    if dry_run:
        result.update(status="DRY", parts=urls)
        return result
    tmp_dir = os.path.join(OUT_DIR, ".parts")
    os.makedirs(tmp_dir, exist_ok=True)
    total = 0
    for g, u in zip(rec["guids"], urls):
        pf = os.path.join(tmp_dir, f"{rec['page_id']}_{g}.mp4")
        info = head_info(u)
        if not info:
            result.update(status="HEAD_FAIL", url=u)
            return result
        size, etag = info
        total += size
        if not (os.path.exists(pf) and os.path.getsize(pf) == size):
            print(f"  بخش {g}: {size/1e6:.0f}MB")
            download_parallel(u, pf, workers=6, block_mb=4, progress=False)
            if not (os.path.exists(pf) and os.path.getsize(pf) == size):
                result.update(status="PART_INCOMPLETE", part=u)
                return result
        part_files.append(pf)
    # concat
    if os.path.exists(dest) and os.path.getsize(dest) == total:
        result.update(status="OK_EXISTS", total_mb=round(total / 1e6, 1))
        return result
    print(f"  concat {len(part_files)} بخش -> {dest}")
    try:
        ffmpeg_concat(part_files, dest)
    except subprocess.CalledProcessError as e:
        result.update(status="CONCAT_FAIL", err=e.stderr.decode()[-300:])
        return result
    # حذف قطعات
    for pf in part_files:
        try:
            os.remove(pf)
        except OSError:
            pass
    if os.path.exists(dest) and os.path.getsize(dest) >= total * 0.97:
        result.update(status="OK", total_mb=round(total / 1e6, 1))
    else:
        result.update(status="CONCAT_BAD",
                      expect=total, got=os.path.getsize(dest))
    return result


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", nargs="*", type=int,
                    help="فقط موقعیت‌های انتخابی (1-5)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)
    todo = [r for r in RECORDS if not args.only or r["sit"] in args.only]

    summary = {}
    for rec in todo:
        label = f"[موقعیت {rec['sit']} #{rec['num']}] {rec['title'][:50]}"
        print(f"\n{'='*60}\n{label}")
        try:
            res = process_one(rec, args.dry_run)
        except Exception as e:
            res = dict(page_id=rec["page_id"], sit=rec["sit"],
                       num=rec["num"], status="ERROR", err=str(e)[:300])
        res["title"] = rec["title"]
        summary[f"Sit{rec['sit']}-{rec['num']}"] = res
        print(f"  => {res['status']}" +
              (f"  {res.get('size_mb') or res.get('total_mb')}MB"
               if res.get("size_mb") or res.get("total_mb") else ""))

    print(f"\n{'='*60}\nخلاصه:")
    for k, v in summary.items():
        print(f"  {k}: {v['status']}  {v.get('title','')[:40]}")
    with open(os.path.join(OUT_DIR, "batch_status.json"), "w",
              encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
