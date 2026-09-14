# -*- coding: utf-8 -*-
"""
sync_photos.py — 把 Google Drive 公開資料夾裡的「全部照片」自動抓進五個課程網站的照片輪播

Drive 裡放什麼就輪播什麼(不去重、不過濾),要拿掉的照片直接在 Drive 刪掉再重跑。

流程:
  1. 讀 Drive 資料夾(含所有子資料夾)的檔案清單
  2. 只下載還沒下載過的照片到 專案/.photo_cache/(Drive 刪掉的也跟著移除;影片、非圖片一律跳過)
  3. 每張:依 EXIF 轉正 → 置中裁切 → 壓縮成 JPG
       - AI 升級計劃(course-upgrade):六角形比例 640×740
       - 其他四站(course / course-elite / course-enterprise / course-tour):正方形 640×640
  4. 依子資料夾輪流排序(輪播才不會連續同一場),輸出 p01.jpg、p02.jpg…
  5. 每站產 assets/photos/manifest.json,網頁載入時自動讀清單長出輪播

用法(在專案根目錄或任何地方執行都可以):
    python "專案/sync_photos.py"                 抓新照片 + 重新裁切全部
    python "專案/sync_photos.py" --no-download   不連網,只用快取重新裁切

需要一次性安裝:
    pip install gdown pillow          (gdown 6 以上)
    (若 Drive 裡有 iPhone 的 HEIC 檔,再加 pip install pillow-heif)

跑完後 git commit + push,Cloudflare 會自動上線。
"""
import argparse
import json
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

try:
    import gdown
except ImportError:
    sys.exit("缺少 gdown,請先執行:pip install gdown")
from PIL import Image, ImageOps

try:  # iPhone HEIC 支援(選配)
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_OK = True
except ImportError:
    HEIC_OK = False

# ───────────────────────────── 設定 ─────────────────────────────
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1GhdzqTVQMvumURYvx6RYwHkhk8xroqw6"

HERE = Path(__file__).resolve().parent                 # 專案/
ROOT = HERE.parent                                     # 政府補助一站式網站
CACHE = HERE / ".photo_cache"                          # 原檔快取(已加進 .gitignore)

HEX = (640, 740)      # 六角形比例 0.866:1,略留餘裕
SQUARE = (640, 640)   # 正方形方塊
TARGETS = [
    (HERE / "5_AI 升級計劃" / "2. 產出" / "assets" / "photos", HEX),
    (ROOT / "course-upgrade" / "assets" / "photos", HEX),
    (ROOT / "course" / "assets" / "photos", SQUARE),
    (ROOT / "course-elite" / "assets" / "photos", SQUARE),
    (ROOT / "course-enterprise" / "assets" / "photos", SQUARE),
    (ROOT / "course-tour" / "assets" / "photos", SQUARE),
]

QUALITY = 78
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"}
# ────────────────────────────────────────────────────────────────


def log(msg):
    print(msg, flush=True)


def list_drive_files():
    """用 gdown 讀取資料夾樹(skip_download 只拿清單,不下載)"""
    log("讀取 Drive 資料夾清單 …")
    files = gdown.download_folder(
        url=DRIVE_FOLDER_URL, output=str(CACHE), quiet=True,
        use_cookies=False, skip_download=True,
    )
    if not files:
        sys.exit("讀不到資料夾內容。請確認資料夾已設成「知道連結的任何人都可檢視」。")
    return files  # 每個有 .id / .path / .local_path


def fetch_file(file_id: str, dest: Path):
    """用 Drive 公開直連下載單一檔案(gdown 逐檔下載常被 Google 擋,改走這條)"""
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    for attempt in range(3):
        with urllib.request.urlopen(req, timeout=120) as r:
            ctype = r.headers.get("Content-Type", "")
            data = r.read()
        if not ctype.startswith("text/html"):
            dest.write_bytes(data)
            return
        time.sleep(3 * (attempt + 1))
    # 直連拿到的是 HTML(大檔病毒掃描確認頁或被擋),最後再交給 gdown 試一次
    gdown.download(id=file_id, output=str(dest), quiet=True, use_cookies=False)


def download_new(files):
    """只下載快取裡還沒有的照片;Drive 已刪掉的,快取也一併刪掉"""
    keep = {Path(f.local_path).resolve() for f in files}
    for p in CACHE.rglob("*"):
        if p.is_file() and p.resolve() not in keep:
            p.unlink()
            log(f"  Drive 已刪除,移除快取 {p.relative_to(CACHE)}")
    todo = [f for f in files
            if Path(f.local_path).suffix.lower() in IMG_EXT and not Path(f.local_path).exists()]
    log(f"Drive 共 {len(files)} 個檔案,照片 {sum(1 for f in files if Path(f.local_path).suffix.lower() in IMG_EXT)} 張,"
        f"需要新下載 {len(todo)} 張")
    for i, f in enumerate(todo, 1):
        Path(f.local_path).parent.mkdir(parents=True, exist_ok=True)
        log(f"  [{i}/{len(todo)}] 下載 {Path(f.local_path).relative_to(CACHE)}")
        fetch_file(f.id, Path(f.local_path))


def collect_cached_photos():
    """從快取收集全部照片,依子資料夾輪流排序"""
    by_folder = {}
    for p in sorted(CACHE.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in IMG_EXT:
            continue
        if p.suffix.lower() in {".heic", ".heif"} and not HEIC_OK:
            log(f"  略過 {p.name}(HEIC 需要 pip install pillow-heif)")
            continue
        by_folder.setdefault(p.parent, []).append(p)

    # 輪流從每個資料夾各取一張,輪播才不會一整段都是同一場
    ordered = []
    queues = [sorted(v) for _, v in sorted(by_folder.items())]
    while any(queues):
        for q in queues:
            if q:
                ordered.append(q.pop(0))
    log(f"照片共 {len(ordered)} 張")
    return ordered


def crop_to(src_im: Image.Image, size) -> Image.Image:
    W, H = size
    w, h = src_im.size
    target = W / H
    if w / h > target:
        cw, ch = int(h * target), h
    else:
        cw, ch = w, int(w / target)
    x0, y0 = (w - cw) // 2, (h - ch) // 2
    return src_im.crop((x0, y0, x0 + cw, y0 + ch)).resize((W, H), Image.LANCZOS)


def write_outputs(photos):
    names = [f"p{i:02d}.jpg" for i in range(1, len(photos) + 1)]
    manifest = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": DRIVE_FOLDER_URL,
        "count": len(names),
        "photos": names,
    }
    # 先把原檔轉正一次,再依各站需要的比例裁切
    opened = [ImageOps.exif_transpose(Image.open(src)).convert("RGB") for src in photos]
    sizes = sorted({size for _, size in TARGETS})
    rendered = {size: [(n, crop_to(im, size)) for n, im in zip(names, opened)] for size in sizes}

    for out, size in TARGETS:
        out.mkdir(parents=True, exist_ok=True)
        for old in out.glob("p*.jpg"):      # 清掉舊的,避免留下已刪除的照片
            old.unlink()
        total = 0
        for name, im in rendered[size]:
            path = out / name
            im.save(path, "JPEG", quality=QUALITY, optimize=True, progressive=True)
            total += path.stat().st_size
        (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"寫入 {out.relative_to(ROOT)}  ({len(names)} 張 · {size[0]}×{size[1]} · {total // 1024} KB)")
    log("manifest.json 已更新")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true", help="不連網,只用快取重新裁切")
    args = ap.parse_args()

    CACHE.mkdir(parents=True, exist_ok=True)
    if not args.no_download:
        download_new(list_drive_files())
    photos = collect_cached_photos()
    if not photos:
        sys.exit("沒有任何照片可用。")
    write_outputs(photos)
    log("完成。接著 git add / commit / push 就會上線。")


if __name__ == "__main__":
    main()
