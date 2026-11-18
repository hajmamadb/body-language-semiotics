# -*- coding: utf-8 -*-
"""
zbd_video.py — ابزار ویدیویی مقالهٔ اول: نشانه‌شناسی زبان بدن در موقعیت‌ها
=========================================================================
امکانات:
1) inventory  : فهرست‌برداری ویدیوها (مدت، ابعاد، نرخ فریم) -> xlsx + json
2) extract    : برداشت فریم‌های معین (لیست زمان یا بازه) -> PNG
3) contact    : ساخت شیت تماس (contact sheet) با برچسب زمان برای کدگذاری
4) sample     : برداشت یکنواخت N فریم از کل ویدیو (نمای کلی)

فرض: ویدیوها در article1/data/ ویدیوها قرار دارند و خروجی‌ها در
article1/output/ برمی‌گردند. استفاده از CLI:

    python zbd_video.py inventory
    python zbd_video.py extract --video "فایلم.mp4" --times "0:02,0:05,1:10"
    python zbd_video.py extract --video "فایلم.mp4" --range "0:02-0:30" --step 2
    python zbd_video.py contact --video "فایلم.mp4" --times "..." [--cols 4]
    python zbd_video.py sample --video "فایلم.mp4" --n 24
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from zbd_common import A1_DATA, A1_OUT, now_ir, safe_name, sec_to_hms, write_json

VIDEO_EXTS = {".mp4", ".mov", ".mkv", ".avi", ".wmv", ".flv", ".m4v", ".webm", ".mpg", ".mpeg"}


def find_videos(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.rglob("*")
        if p.suffix.lower() in VIDEO_EXTS and p.is_file()
    )


# ------------------------------------------------------------------ inventory
def inventory(vid_dir: Path | None = None) -> Path:
    """فهرست‌برداری همهٔ ویدیوها در article1/data/ویدیوها."""
    import cv2
    vid_dir = Path(vid_dir) if vid_dir else A1_DATA / "ویدیوها"
    out_json = A1_OUT / "video_inventory.json"
    out_xlsx = A1_OUT / "فهرست_ویدیوها.xlsx"

    rows = []
    for vp in find_videos(vid_dir):
        cap = cv2.VideoCapture(str(vp))
        if not cap.isOpened():
            rows.append({"file": str(vp), "error": "not opened"})
            continue
        fps = cap.get(cv2.CAP_PROP_FPS) or 0
        n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
        w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        duration = n / fps if fps else 0
        rel = vp.relative_to(A1_DATA)
        rows.append({
            "file": str(rel), "abs": str(vp),
            "width": w, "height": h, "fps": round(fps, 3),
            "frames": n, "duration_sec": round(duration, 2),
            "duration_hms": sec_to_hms(duration),
        })
        cap.release()

    write_json(out_json, {"generated_at": now_ir(), "videos": rows})
    # xlsx سادهٔ RTL
    from zbd_sheet import SheetBuilder
    sb = SheetBuilder("فهرست ویدیوها")
    headers = ["فایل (نسبی)", "عرض", "ارتفاع", "FPS", "تعداد فریم",
               "مدت (ثانیه)", "مدت (h:m:s)"]
    sb.set_columns(headers, [46, 10, 10, 10, 12, 12, 14])
    for r in rows:
        sb.add_row(r.get("file"), r.get("width"), r.get("height"), r.get("fps"),
                   r.get("frames"), r.get("duration_sec"), r.get("duration_hms"))
    sb.save(out_xlsx)
    return out_xlsx


# ------------------------------------------------------------------ helpers
def _open_cap(path: Path):
    import cv2
    cap = cv2.VideoCapture(str(path))
    if not cap.isOpened():
        raise SystemExit(f"نمی‌توان ویدیو را باز کرد: {path}")
    return cap


def _grab_frame(cap, t_sec: float):
    import cv2
    fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    cap.set(cv2.CAP_PROP_POS_MSEC, t_sec * 1000.0)
    ok, frame = cap.read()
    if not ok:  # بازگشت به نزدیک‌ترین فریم معتبر
        target = int(t_sec * fps)
        cap.set(cv2.CAP_PROP_POS_FRAMES, max(0, target))
        ok, frame = cap.read()
    return frame if ok else None


def _imwrite_unicode(path: Path, frame, ext: str = ".png") -> bool:
    """cv2.imwrite در ویندوز با مسیر فارسی کار نمی‌کند؛ با imencode + open حل می‌شود."""
    import cv2
    ok, buf = cv2.imencode(ext, frame)
    if not ok:
        return False
    Path(path).write_bytes(buf.tobytes())
    return True


def _parse_times(spec: str) -> list[float]:
    """لیست زمان‌ها: '0:02,0:05,1:10' یا '12.5,30'."""
    from zbd_common import hms_to_sec
    return [hms_to_sec(p) for p in spec.split(",") if p.strip()]


def _parse_range(spec: str, step: float) -> list[float]:
    a, b = [s.strip() for s in spec.split("-")]
    from zbd_common import hms_to_sec
    start, end = hms_to_sec(a), hms_to_sec(b)
    if end < start:
        start, end = end, start
    out = []
    t = start
    while t <= end + 1e-6:
        out.append(t)
        t += step
    return out


# ------------------------------------------------------------------ extract
def extract(video: Path, times: list[float], out_dir: Path) -> list[Path]:
    import cv2
    out_dir.mkdir(parents=True, exist_ok=True)
    cap = _open_cap(video)
    saved: list[Path] = []
    for t in times:
        frame = _grab_frame(cap, t)
        if frame is None:
            print(f"  ! فریم {sec_to_hms(t)} یافت نشد")
            continue
        name = f"{safe_name(video.stem)}_{sec_to_hms(t).replace(':', '_')}.png"
        p = out_dir / name
        if _imwrite_unicode(p, frame, ".png"):
            saved.append(p)
        else:
            print(f"  ! نوشتن {p} ناموفق بود")
    cap.release()
    return saved


# ------------------------------------------------------------------ contact
def contact_sheet(video: Path, times: list[float], out_path: Path,
                  cols: int = 4, thumb_w: int = 360) -> Path:
    """شیت تماس: شبکه‌ای از فریم‌ها با برچسب زمان روی هر تصویر."""
    import cv2
    from PIL import Image, ImageDraw, ImageFont
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cap = _open_cap(video)

    thumbs = []
    for t in times:
        frame = _grab_frame(cap, t)
        if frame is None:
            continue
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        ratio = thumb_w / img.width
        img = img.resize((thumb_w, max(1, int(img.height * ratio))))
        # حاشیهٔ زیر برای برچسب زمان
        label_h = 26
        canvas = Image.new("RGB", (thumb_w, img.height + label_h), "black")
        canvas.paste(img, (0, 0))
        d = ImageDraw.Draw(canvas)
        d.text((4, img.height + 4), sec_to_hms(t), fill="yellow")
        thumbs.append(canvas)
    cap.release()

    if not thumbs:
        raise SystemExit("هیچ فریمی برای شیت تماس یافت نشد.")

    th_h = thumbs[0].height
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * thumb_w, rows * th_h), "black")
    for i, th in enumerate(thumbs):
        sheet.paste(th, ((i % cols) * thumb_w, (i // cols) * th_h))
    sheet.save(out_path, quality=92)
    return out_path


# ------------------------------------------------------------------ sample
def sample_frames(video: Path, n: int, out_dir: Path) -> list[float]:
    import cv2
    cap = _open_cap(video)
    fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    total = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
    dur = total / fps if fps else 0
    cap.release()
    if dur <= 0:
        raise SystemExit("مدت ویدیو صفر است؛ نمی‌توان نمونه‌برداری کرد.")
    n = max(2, int(n))
    times = [dur * i / (n - 1) for i in range(n)]
    times = [t for t in times if t < dur - 0.05]
    extract(video, times, out_dir)
    return times


# ------------------------------------------------------------------ CLI
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("inventory", help="فهرست‌برداری ویدیوها")

    pe = sub.add_parser("extract", help="برداشت فریم‌های معین")
    pe.add_argument("--video", required=True)
    pe.add_argument("--times", help="لیست زمان: 0:02,0:05 یا ثانیه")
    pe.add_argument("--range", dest="rng", help="بازه: 0:02-0:30")
    pe.add_argument("--step", type=float, default=5.0, help="فاصلهٔ برداشت در بازه (ثانیه)")
    pe.add_argument("--out", help="پوشهٔ خروجی (پیش‌فرض: article1/output/فریم‌ها)")

    pc = sub.add_parser("contact", help="ساخت شیت تماس")
    pc.add_argument("--video", required=True)
    pc.add_argument("--times", required=True)
    pc.add_argument("--cols", type=int, default=4)
    pc.add_argument("--out", help="مسیر خروجی png (پیش‌فرض: article1/output)")

    ps = sub.add_parser("sample", help="نمونه‌برداری یکنواخت")
    ps.add_argument("--video", required=True)
    ps.add_argument("--n", type=int, default=24)
    ps.add_argument("--out", help="پوشهٔ خروجی")

    args = ap.parse_args(argv)

    if args.cmd == "inventory":
        xlsx = inventory()
        print("فهرست ویدیوها:", xlsx)
        print(" JSON:", A1_OUT / "video_inventory.json")
        return 0

    v = Path(args.video)
    if not v.is_absolute():
        # ترتیب جستجو: مسیر داده‌شده -> زیرپوشهٔ ویدیوها -> ریشهٔ data
        for cand in (A1_DATA / v, A1_DATA / "ویدیوها" / v, Path(v)):
            if cand.exists():
                v = cand
                break
    if not v.exists():
        print(f"ویدیو یافت نشد: {v} (آیا در article1/data/ویدیوها قرار دارد؟)")
        return 1

    if args.cmd == "extract":
        if not args.times and not args.rng:
            print("یا --times یا --range لازم است")
            return 1
        times = []
        if args.times:
            times += _parse_times(args.times)
        if args.rng:
            times += _parse_range(args.rng, args.step)
        times = sorted(set(times))
        out_dir = Path(args.out) if args.out else A1_OUT / "فریم‌ها" / safe_name(v.stem)
        saved = extract(v, times, out_dir)
        print(f"{len(saved)} فریم در: {out_dir}")
        return 0

    if args.cmd == "contact":
        times = _parse_times(args.times)
        out = Path(args.out) if args.out else A1_OUT / f"contact_{safe_name(v.stem)}.png"
        p = contact_sheet(v, times, out, cols=args.cols)
        print("شیت تماس:", p)
        return 0

    if args.cmd == "sample":
        out_dir = Path(args.out) if args.out else A1_OUT / "نمونه‌ها" / safe_name(v.stem)
        times = sample_frames(v, args.n, out_dir)
        print(f"{len(times)} فریم نمونه در: {out_dir}")
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
