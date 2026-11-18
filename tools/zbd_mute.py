# -*- coding: utf-8 -*-
"""
zbd_mute.py — صامت‌سازی ویدیوها (حذف صدا) با ffmpeg
=====================================================
بر اساس اصل روش‌شناختی «صامت‌سازی»: پیش از تحلیل، صدای ویدیوی مبنا
کاملا قطع می‌شود تا ریتم، لحن و مفاهیم سخن موجب سوگیری نشود.

کاربرد:
    python zbd_mute.py [--video "فایل.mp4"] [--all] [--out "پوشه-خروجی"]

پیش‌فرض: صامت‌سازی همه ویدیوها در article1/data/ویدیوها/
خروجی: article1/data/صامت/ (فایل‌های صامت)
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zbd_common import A1_DATA, now_ir, write_json

FFMPEG = r"C:\Users\MamaDDelta\AppData\Roaming\Python\Python315\site-packages\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v9.0.2.exe"

VIDEO_EXTS = {".mp4", ".mov", ".mkv", ".avi", ".wmv", ".flv", ".m4v", ".webm", ".mpg", ".mpeg"}


def find_videos(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.rglob("*")
        if p.suffix.lower() in VIDEO_EXTS and p.is_file()
    )


def mute_one(src: Path, out_dir: Path) -> dict:
    """صامت‌سازی یک ویدیو: حذف جریان صوتی با stream copy."""
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / src.name
    if dst.exists() and dst.stat().st_size > 1000:
        return {"file": src.name, "status": "SKIP_EXISTS"}
    # Delete 0-byte partial
    if dst.exists():
        dst.unlink()

    cmd = [
        FFMPEG, "-y",
        "-i", str(src),
        "-an",                          # حذف صدا
        "-c:v", "copy",                 # کپی جریان ویدیو (بدون ریکود)
        "-movflags", "+faststart",
        str(dst),
    ]
    # Two-pass: mute to C: temp first (D: may be full), then swap
    C_TEMP = Path(r"C:\Users\MamaDDelta\AppData\Local\Temp\zbd_mute_tmp")
    C_TEMP.mkdir(parents=True, exist_ok=True)
    tmp_dst = C_TEMP / src.name

    cmd = [
        FFMPEG, "-y",
        "-i", str(src),
        "-an",                          # حذف صدا
        "-c:v", "copy",                 # کپی جریان ویدیو (بدون ریکود)
        "-movflags", "+faststart",
        str(tmp_dst),
    ]
    result = subprocess.run(cmd, capture_output=True)
    if result.returncode != 0:
        err = result.stderr.decode(errors="replace")[-300:]
        tmp_dst.unlink(missing_ok=True)
        return {"file": src.name, "status": "ERROR", "err": err}

    # Swap: move original aside, move muted in, delete original
    src_mb = src.stat().st_size / 1e6
    orig_backup = C_TEMP / (src.name + ".orig")
    shutil.move(str(src), str(orig_backup))
    shutil.move(str(tmp_dst), str(dst))
    orig_backup.unlink(missing_ok=True)

    dst_mb = dst.stat().st_size / 1e6
    saved_mb = src_mb - dst_mb
    return {
        "file": src.name,
        "status": "OK",
        "src_mb": round(src_mb, 1),
        "dst_mb": round(dst_mb, 1),
        "saved_mb": round(saved_mb, 1),
        "saved_pct": f"{saved_mb/src_mb*100:.1f}%" if src_mb else "0%",
    }


def main():
    ap = argparse.ArgumentParser(description="صامت‌سازی ویدیوها")
    ap.add_argument("--video", help="فایل ویدیوی مشخص")
    ap.add_argument("--all", action="store_true", help="همه ویدیوها")
    ap.add_argument("--out", help="پوشه خروجی (پیش‌فرض: article1/data/صامت)")
    args = ap.parse_args()

    vid_dir = A1_DATA / "ویدیوها"
    out_dir = Path(args.out) if args.out else A1_DATA / "صامت"

    if args.video:
        v = Path(args.video)
        if not v.is_absolute():
            for cand in (vid_dir / v, v):
                if cand.exists():
                    v = cand
                    break
        if not v.exists():
            print(f"ویدیو یافت نشد: {v}")
            return 1
        videos = [v]
    else:
        videos = find_videos(vid_dir)

    if not videos:
        print("ویدیویی یافت نشد.")
        return 1

    print(f"صامت‌سازی {len(videos)} ویدیو ...")
    results = []
    for vp in videos:
        print(f"  {vp.name} ... ", end="", flush=True)
        r = mute_one(vp, out_dir)
        if r["status"] == "OK":
            print(f"OK  {r['src_mb']}MB -> {r['dst_mb']}MB  (حذف {r['saved_mb']}MB)")
        elif r["status"] == "SKIP_EXISTS":
            print("SKIP (موجود است)")
        else:
            print(f"ERROR: {r.get('err', '')[:100]}")
        results.append(r)

    # خلاصه
    ok = [r for r in results if r["status"] == "OK"]
    err = [r for r in results if r["status"] == "ERROR"]
    total_saved = sum(r.get("saved_mb", 0) for r in ok)
    print(f"\nخلاصه: {len(ok)} موفق, {len(err)} خطا, کل ذخیره‌شده: {total_saved:.0f}MB")

    # JSON
    out_json = out_dir / "mute_report.json"
    write_json(out_json, {"generated_at": now_ir(), "results": results})
    print(f"گزارش: {out_json}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
