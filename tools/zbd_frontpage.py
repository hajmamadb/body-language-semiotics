# -*- coding: utf-8 -*-
"""
zbd_frontpage.py — ابزار تصویر مقالهٔ دوم: بازنمایی زبان بدن در صفحهٔ اول
=========================================================================
داده‌ها در article2/data/صفحه‌اول (عکس کامل صفحه) و article2/data/برش‌ها
(برش‌های عکس از صفحه) قرار می‌گیرند؛ خروجی‌ها در article2/output/.

امکانات:
1) inventory : فهرست تصاویر صفحهٔ اول و برش‌ها (ابعاد، DPI تخمینی) -> xlsx
2) grid      : ترسیم شبکهٔ مرجع ۳×۳ روی تصویر صفحه (برای کدگذاری جایگاه)
3) crop      : برش یک ناحیه از صفحه -> png (برای تحلیل دقیق نشانهٔ بدنی)
4) grid3     : شیت مقایسهٔ ۳ صفحهٔ یک رویداد کنار هم (رویدادهای مشترک)
5) contact   : شیت تماس از همهٔ برش‌ها برای کدگذاری سریع

CLI نمونه:
    python zbd_frontpage.py inventory
    python zbd_frontpage.py grid   --image "روزA_1403-01-10.png"
    python zbd_frontpage.py crop   --image "روزA_1403-01-10.png" --box "0.1,0.2,0.6,0.7"
    python zbd_frontpage.py grid3  --images "روزA_1403-01-10.png روزB_1403-01-10.png روزC_1403-01-10.png"
    python zbd_frontpage.py contact --from-crops
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from zbd_common import A2_DATA, A2_OUT, now_ir, safe_name, write_json

IMG_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".bmp"}


def find_images(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.rglob("*")
        if p.suffix.lower() in IMG_EXTS and p.is_file()
    )


def _resolve(data_arg: str | Path) -> Path:
    p = Path(data_arg)
    if p.is_absolute() and p.exists():
        return p
    for cand in (A2_DATA / p, A2_DATA / "صفحه‌اول" / p, A2_DATA / "برش‌ها" / p):
        if cand.exists():
            return cand
    raise SystemExit(f"تصویر یافت نشد: {data_arg}")


# ------------------------------------------------------------------ inventory
def inventory() -> Path:
    from PIL import Image
    rows = []
    for kind, folder in (("صفحهٔ کامل", A2_DATA / "صفحه‌اول"), ("برش", A2_DATA / "برش‌ها")):
        if not folder.exists():
            continue
        for p in find_images(folder):
            try:
                with Image.open(p) as im:
                    w, h = im.size
                    dpi = im.info.get("dpi")
                    dpi = tuple(round(d, 1) for d in dpi) if dpi else None
            except Exception as e:  # noqa: BLE001
                rows.append({"type": kind, "file": str(p), "error": str(e)})
                continue
            rows.append({
                "type": kind, "file": str(p.relative_to(A2_DATA)),
                "width": w, "height": h,
                "aspect": round(w / h, 3) if h else None,
                "dpi": dpi,
            })
    out_json = A2_OUT / "image_inventory.json"
    write_json(out_json, {"generated_at": now_ir(), "images": rows})

    from zbd_sheet import SheetBuilder
    sb = SheetBuilder("فهرست تصاویر")
    sb.set_columns(["نوع", "فایل", "عرض", "ارتفاع", "نسبت عرض/ارتفاع", "DPI"],
                   [14, 50, 10, 10, 16, 14])
    for r in rows:
        dpi = r.get("dpi")
        sb.add_row(r.get("type"), r.get("file"), r.get("width"), r.get("height"),
                   r.get("aspect"), (f"{dpi[0]}/{dpi[1]}" if dpi else ""))
    out_xlsx = A2_OUT / "فهرست_تصاویر.xlsx"
    sb.save(out_xlsx)
    return out_xlsx


# ------------------------------------------------------------------ grid
def grid(image: Path, out: Path | None = None) -> Path:
    """ترسیم شبکهٔ ۳×۳ مرجع + خطوط یک‌سوم روی تصویر صفحه."""
    from PIL import Image, ImageDraw
    image = _resolve(image)
    out = Path(out) if out else A2_OUT / f"grid_{image.stem}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(image).convert("RGB") as im:
        w, h = im.size
        d = ImageDraw.Draw(im)
        lw = max(2, w // 500)
        for x in (w // 3, 2 * w // 3):
            d.line([(x, 0), (x, h)], fill="red", width=lw)
        for y in (h // 3, 2 * h // 3):
            d.line([(0, y), (w, y)], fill="red", width=lw)
        # برچسب ناحیه‌ها (راست‌به‌چپ: بالا-راست ...)
        labels = [
            (0.06, 0.04, "راست-بالا"), (0.42, 0.04, "وسط-بالا"), (0.80, 0.04, "چپ-بالا"),
            (0.06, 0.36, "راست-میانه"), (0.42, 0.36, "وسط-میانه"), (0.80, 0.36, "چپ-میانه"),
            (0.06, 0.90, "راست-پایین"), (0.42, 0.90, "وسط-پایین"), (0.80, 0.90, "چپ-پایین"),
        ]
        fs = max(14, w // 60)
        for fx, fy, txt in labels:
            d.text((int(fx * w), int(fy * h)), txt, fill="yellow")
        im.save(out, quality=92)
    return out


# ------------------------------------------------------------------ crop
def crop(image: Path, box: str, out: Path | None = None) -> Path:
    """برش ناحیه با مختصات نسبی 'x1,y1,x2,y2' (هر کدوم ۰ تا ۱)."""
    from PIL import Image
    image = _resolve(image)
    try:
        x1, y1, x2, y2 = (float(v.strip()) for v in box.split(","))
    except ValueError:
        raise SystemExit("box باید به صورت x1,y1,x2,y2 با اعداد ۰ تا ۱ باشد")
    if not all(0 <= v <= 1 for v in (x1, y1, x2, y2)) or x1 >= x2 or y1 >= y2:
        raise SystemExit("box نامعتبر است")
    out = Path(out) if out else A2_OUT / "برش‌ها" / f"crop_{image.stem}_{safe_name(box.replace(',', '_'))}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(image).convert("RGB") as im:
        w, h = im.size
        c = im.crop((int(x1 * w), int(y1 * h), int(x2 * w), int(y2 * h)))
        c.save(out, quality=95)
    return out


# ------------------------------------------------------------------ grid3 (مقایسهٔ رویداد مشترک)
def grid3(images: list[str], out: Path | None = None, label_h: int = 40) -> Path:
    """چینش افقی چند صفحهٔ اول (پوشش یک رویداد مشترک) با عنوان هر روزنامه."""
    from PIL import Image, ImageDraw
    resolved = [_resolve(i) for i in images]
    out = Path(out) if out else A2_OUT / "مقایسه" / f"compare_{len(resolved)}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    thumbs = []
    max_h = 900
    for p in resolved:
        with Image.open(p).convert("RGB") as im:
            ratio = max_h / im.height
            im = im.resize((max(1, int(im.width * ratio)), max_h))
            canvas = Image.new("RGB", (im.width + 20, im.height + label_h), (40, 40, 40))
            canvas.paste(im, (10, label_h))
            d = ImageDraw.Draw(canvas)
            d.text((10, 8), p.stem, fill="white")
            thumbs.append(canvas)
    tw = sum(t.width for t in thumbs)
    th = max(t.height for t in thumbs)
    sheet = Image.new("RGB", (tw, th), (20, 20, 20))
    x = 0
    for t in thumbs:
        sheet.paste(t, (x, 0))
        x += t.width
    sheet.save(out, quality=90)
    return out


# ------------------------------------------------------------------ contact (برش‌ها)
def contact_crops(cols: int = 4, thumb_w: int = 420) -> Path:
    """شیت تماس از همهٔ برش‌ها در article2/data/برش‌ها."""
    from PIL import Image, ImageDraw
    folder = A2_DATA / "برش‌ها"
    if not folder.exists():
        raise SystemExit("پوشهٔ article2/data/برش‌ها وجود ندارد")
    crops = find_images(folder)
    if not crops:
        raise SystemExit("برشی در article2/data/برش‌ها یافت نشد")
    thumbs = []
    for p in crops:
        with Image.open(p).convert("RGB") as im:
            ratio = thumb_w / im.width
            im = im.resize((thumb_w, max(1, int(im.height * ratio))))
            lab = 30
            canvas = Image.new("RGB", (thumb_w, im.height + lab), (30, 30, 30))
            canvas.paste(im, (0, 0))
            d = ImageDraw.Draw(canvas)
            d.text((4, im.height + 6), p.stem[:46], fill="yellow")
            thumbs.append(canvas)
    th_h = thumbs[0].height
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * thumb_w, rows * th_h), "black")
    for i, t in enumerate(thumbs):
        sheet.paste(t, ((i % cols) * thumb_w, (i // cols) * th_h))
    out = A2_OUT / "contact_برش‌ها.png"
    sheet.save(out, quality=90)
    return out


# ------------------------------------------------------------------ CLI
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("inventory")

    pg = sub.add_parser("grid")
    pg.add_argument("--image", required=True)
    pg.add_argument("--out")

    pc = sub.add_parser("crop")
    pc.add_argument("--image", required=True)
    pc.add_argument("--box", required=True, help="x1,y1,x2,y2 (نسبی ۰ تا ۱)")
    pc.add_argument("--out")

    p3 = sub.add_parser("grid3")
    p3.add_argument("--images", required=True, help="فاصله‌دار، صفحه‌های یک رویداد مشترک")
    p3.add_argument("--out")

    pk = sub.add_parser("contact")
    pk.add_argument("--cols", type=int, default=4)

    args = ap.parse_args(argv)
    if args.cmd == "inventory":
        print("فهرست تصاویر:", inventory())
        print(" JSON:", A2_OUT / "image_inventory.json")
        return 0
    if args.cmd == "grid":
        print("شبکهٔ مرجع:", grid(args.image, args.out))
        return 0
    if args.cmd == "crop":
        print("برش:", crop(args.image, args.box, args.out))
        return 0
    if args.cmd == "grid3":
        print("شیت مقایسه:", grid3(args.images.split(), args.out))
        return 0
    if args.cmd == "contact":
        print("شیت تماس برش‌ها:", contact_crops(cols=args.cols))
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
