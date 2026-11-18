# -*- coding: utf-8 -*-
"""
zbd_vision.py — استخراج ویژگی‌های بدنی با مدل بینایی محلی (YOLOv8)
====================================================================
هدف: برای «دقیق‌تر و متنوع‌تر» کردن کتاب‌کد، از عکس‌ها و فریم‌های ویدیو
ویژگی‌های عددی قابل‌اندازه‌گیری استخراج می‌شود (بدون نیاز به مدل چندحسی
زبان‌تصویری). خروجی یک JSON است که برای بازنگری کدها استفاده می‌شود.

امکانات:
- اسکن ۱۰ عکس در پوشهٔ عکس/ (تشخیص انسان/صورت/دست)
- برداشت ۱۲ فریم از هر ویدیو در ویدئو/ + اسکن
- محاسبهٔ نسبت‌های هندسی (aspect, نسبت ارتفاع صورت به بدن، ...)

CLI نمونه:
    python zbd_vision.py scan --photos
    python zbd_vision.py scan --videos
    python zbd_vision.py scan --all
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from zbd_common import A1_DATA, A1_OUT, ROOT, now_ir, sec_to_hms, write_json

PHOTO_DIR = ROOT / "عکس"
VIDEO_DIR = ROOT / "ویدئو"
N_VIDEO_FRAMES = 12


def _load_yolo():
    from ultralytics import YOLO
    model = YOLO("yolov8n.pt")  # نسخهٔ نانو (سبک و کافی برای این کار)
    return model


def _scan_image(model, img_path: Path, img_idx: int | None = None,
                time_sec: float | None = None) -> dict:
    """اسکن یک تصویر: تشخیص انسان + صورت + دست + ویژگی‌های هندسی."""
    import cv2
    from PIL import Image
    out: dict = {
        "file": img_path.name,
        "index": img_idx,
        "time_sec": round(time_sec, 2) if time_sec is not None else None,
        "time_hms": sec_to_hms(time_sec) if time_sec is not None else None,
    }
    try:
        cap = cv2.VideoCapture(str(img_path))
        if cap.isOpened():
            ok, frame = cap.read()
            cap.release()
        else:
            frame = None
        if frame is None:
            frame = cv2.imdecode(
                (lambda b: b)(img_path.read_bytes()), cv2.IMREAD_COLOR)
        if frame is None:
            out["error"] = "read failed"
            return out
    except Exception as e:  # noqa: BLE001
        out["error"] = f"read error: {e}"
        return out

    h_img, w_img = frame.shape[:2]
    out["width"], out["height"] = w_img, h_img
    out["aspect"] = round(w_img / h_img, 3) if h_img else None

    try:
        results = model(frame, verbose=False)
    except Exception as e:  # noqa: BLE001
        out["error"] = f"model error: {e}"
        return out

    r = results[0]
    boxes, cls_ids = [], []
    if r.boxes is not None:
        for box, cls in zip(r.boxes.xyxy, r.boxes.cls):
            boxes.append([round(float(v), 1) for v in box])
            cls_ids.append(int(cls))
    names = r.names if hasattr(r, "names") else {}
    objs = []
    for (x1, y1, x2, y2), c in zip(boxes, cls_ids):
        label = names.get(c, str(c))
        objs.append({
            "class": label,
            "conf": 0.0,  # confidence not easily per-box in this path
            "box": [x1, y1, x2, y2],
            "w": round(x2 - x1, 1), "h": round(y2 - y1, 1),
            "aspect": round((x2 - x1) / (y2 - y1), 3) if (y2 > y1) else None,
            # موقعیت نسبی در قاب (0-1)
            "cx_rel": round(((x1 + x2) / 2) / w_img, 3),
            "cy_rel": round(((y1 + y2) / 2) / h_img, 3),
            "area_rel": round(((x2 - x1) * (y2 - y1)) / (w_img * h_img), 4),
        })
    out["objects"] = objs

    # شاخص‌های محاسبه‌شده برای کتاب‌کد
    people = [o for o in objs if o["class"] == "person"]
    faces = [o for o in objs if o["class"] == "face"]
    hands = [o for o in objs if o["class"] == "hand"]
    out["counts"] = {"person": len(people), "face": len(faces), "hand": len(hands)}
    if people:
        p = people[0]
        out["primary_person"] = p
        if faces:
            f = faces[0]
            out["face_to_body_h"] = round(f["h"] / p["h"], 3) if p["h"] else None
        # وضعیت عمودی: آیا سر در بالای قاب است؟
        head_top_rel = p["box"][1] / h_img
        out["head_top_rel"] = round(head_top_rel, 3)
    return out


def _scan_video(model, vid_path: Path) -> list[dict]:
    """برداشت ۱۲ فریم یکنواخت + اسکن هرکدام."""
    import cv2
    cap = cv2.VideoCapture(str(vid_path))
    if not cap.isOpened():
        return [{"file": vid_path.name, "error": "cannot open"}]
    fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    dur = n / fps if fps else 0
    times = [dur * i / (N_VIDEO_FRAMES - 1) for i in range(N_VIDEO_FRAMES)]
    out = []
    import tempfile
    for i, t in enumerate(times):
        cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000.0)
        ok, frame = cap.read()
        if not ok:
            continue
        # ذخیرهٔ موقت برای اسکن
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = Path(tmp.name)
            okw, buf = cv2.imencode(".png", frame)
            if okw:
                tmp_path.write_bytes(buf.tobytes())
        try:
            rec = _scan_image(model, tmp_path, img_idx=i, time_sec=t)
            rec["video"] = vid_path.name
            out.append(rec)
        finally:
            tmp_path.unlink(missing_ok=True)
    cap.release()
    return out


def scan_photos() -> Path:
    model = _load_yolo()
    results = []
    photos = sorted(p for p in PHOTO_DIR.iterdir()
                    if p.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".bmp"})
    print(f"اسکن {len(photos)} عکس در {PHOTO_DIR} ...")
    for i, p in enumerate(photos):
        print(f"  [{i+1}/{len(photos)}] {p.name}")
        results.append(_scan_image(model, p, img_idx=i))
    out = A1_OUT / "vision_photos.json"
    write_json(out, {"generated_at": now_ir(), "type": "photos", "records": results})
    print("خروجی:", out)
    return out


def scan_videos() -> Path:
    model = _load_yolo()
    results = []
    videos = sorted(VIDEO_DIR.glob("*.mp4"))
    print(f"اسکن {len(videos)} ویدیو (هرکدام {N_VIDEO_FRAMES} فریم) ...")
    for v in videos:
        print(f"  {v.name}")
        frames = _scan_video(model, v)
        results.extend(frames)
    out = A1_OUT / "vision_videos.json"
    write_json(out, {"generated_at": now_ir(), "type": "videos", "records": results})
    print("خروجی:", out)
    return out


def scan_all() -> tuple[Path, Path]:
    p1 = scan_photos()
    p2 = scan_videos()
    return p1, p2


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    ps = sub.add_parser("scan")
    ps.add_argument("--photos", action="store_true")
    ps.add_argument("--videos", action="store_true")
    ps.add_argument("--all", action="store_true")
    args = ap.parse_args(argv)

    if args.cmd == "scan":
        if args.all or not (args.photos or args.videos):
            scan_all()
        elif args.photos:
            scan_photos()
        elif args.videos:
            scan_videos()
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
