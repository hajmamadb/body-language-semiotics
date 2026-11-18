# -*- coding: utf-8 -*-
"""
zbd_parallel_dl.py — دانلود‌کنندهٔ موازی (parallel range) از CDN کhamenei.ir
============================================================
CDN سایت khamenei.ir محافظت CloudGuard ندارد (فقط HTML صفحه محافظت دارد)
و از Range Request پشتیبانی می‌کند. این ابزار، فایل را در چند بلوک موازی
دانلود و در کنار هم می‌نویسد.

کاربرد:
    python zbd_parallel_dl.py <URL> <output.mp4> [--workers 6] [--block 2MB]

نکته: فقط برای CDN مستقیم (idc0-cdn*.khamenei.ir) است، نه برای صفحهٔ HTML.
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " \
     "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def get_file_size(url: str) -> int:
    """اندازهٔ کل فایل با HEAD (یا Range 0-0)."""
    try:
        r = requests.head(url, headers={"User-Agent": UA}, timeout=20, allow_redirects=True)
        if r.status_code in (200, 206):
            return int(r.headers.get("Content-Length", 0))
    except Exception:
        pass
    # fallback: Range 0-0
    r = requests.get(url, headers={"User-Agent": UA, "Range": "bytes=0-0"},
                     timeout=20, stream=True)
    cr = r.headers.get("Content-Range", "")  # bytes 0-0/TOTAL
    r.close()
    if "/" in cr:
        return int(cr.split("/")[-1])
    raise RuntimeError(f"Could not determine file size: {url}")


def download_block(url: str, start: int, end: int, dest: str, lock) -> tuple[int, int, bool]:
    """دانلود یک بلوک [start, end] به آدرس مشخص در فایل مقصد."""
    pos = start
    deadline = time.time() + 300  # 5 دقیقه برای هر بلوک
    while pos <= end:
        if time.time() > deadline:
            return (start, pos - 1, False)
        try:
            with requests.get(
                url,
                headers={"User-Agent": UA, "Range": f"bytes={pos}-{end}"},
                timeout=60, stream=True,
            ) as r:
                if r.status_code not in (200, 206):
                    return (start, pos - 1, False)
                for chunk in r.iter_content(chunk_size=256 * 1024):
                    if chunk:
                        with lock:
                            with open(dest, "r+b") as f:
                                f.seek(pos)
                                f.write(chunk)
                        pos += len(chunk)
        except Exception:
            # retry
            time.sleep(1)
    return (start, end, True)


def download_parallel(url: str, dest: str, workers: int = 6,
                      block_mb: float = 4.0, progress: bool = True) -> None:
    size = get_file_size(url)
    if size == 0:
        raise RuntimeError("File size is 0 or unknown")
    print(f"  File size: {size / 1e6:.1f} MB")

    # اگر قبلاً کامل دانلود شده، رد شو
    if os.path.exists(dest):
        cur = os.path.getsize(dest)
        if cur >= size:
            print(f"  Already complete ({cur} bytes) — skip")
            return

    block = int(block_mb * 1024 * 1024)
    # بلوک‌بندی
    blocks = []
    start = 0
    while start < size:
        end = min(start + block - 1, size - 1)
        blocks.append((start, end))
        start = end + 1
    print(f"  Blocks: {len(blocks)} × {block_mb:.0f} MB, workers: {workers}")

    # ساخت فایل مقصد (empty)
    dest_dir = os.path.dirname(dest)
    if dest_dir:
        os.makedirs(dest_dir, exist_ok=True)
    with open(dest, "wb") as f:
        f.truncate(size)

    lock = __import__("threading").Lock()
    done_bytes = 0
    t0 = time.time()

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(download_block, url, s, e, dest, lock): (s, e) for s, e in blocks}
        for fut in as_completed(futs):
            s, e = futs[fut]
            try:
                bs, be, ok = fut.result()
                if not ok:
                    print(f"  ⚠ Block [{bs}-{be}] incomplete (got to {be})")
                done_bytes += max(0, min(be, e) - bs + 1)
            except Exception as ex2:
                print(f"  ✗ Block error: {ex2}")
            if progress:
                pct = done_bytes / size * 100
                el = time.time() - t0
                spd = done_bytes / el / 1024 if el > 0 else 0
                eta = (size - done_bytes) / (spd * 1024) if spd > 0 else 0
                print(f"  {pct:5.1f}%  {done_bytes/1e6:8.1f}/{size/1e6:.0f} MB  "
                      f"{spd:6.0f} KB/s  ETA {eta/60:4.0f} min", end="\r")

    # بررسی نهایی
    final = os.path.getsize(dest)
    if final < size:
        missing = size - final
        print(f"\n  ⚠ File incomplete: {final}/{size} ({missing} bytes missing)")
        # تلاش دوباره برای بخش‌های ناقص
        print("  Retrying missing parts...")
        # پیدا کردن بخش‌های خالی
        with open(dest, "rb") as f:
            data = f.read()
        # بخش‌های صفر را پیدا کن
        import re as _re
        # ساده‌سازی: فقط اگر کمتر از 10٪ ناقص است، دوباره کل را دانلود کن
        if missing / size < 0.1:
            # دانلود مجدد کل (ساده)
            download_parallel(url, dest, workers, block_mb, progress)
        else:
            raise RuntimeError(f"File too incomplete: {missing} bytes missing")
    else:
        print(f"\n  ✓ Done: {final} bytes  in {time.time()-t0:.0f}s")


def main() -> None:
    ap = argparse.ArgumentParser(description="Parallel range downloader")
    ap.add_argument("url")
    ap.add_argument("dest")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--block", type=float, default=4.0, help="Block size in MB")
    args = ap.parse_args()
    download_parallel(args.url, args.dest, args.workers, args.block)


if __name__ == "__main__":
    main()
