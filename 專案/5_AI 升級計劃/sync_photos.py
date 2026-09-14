# -*- coding: utf-8 -*-
"""
sync_photos.py — 把 Google Drive 公開資料夾裡的「全部照片」自動抓進網站輪播

流程:
  1. 讀 Drive 資料夾(含所有子資料夾)的檔案清單
  2. 只下載還沒下載過的照片到 .photo_cache/(影片、非圖片一律跳過)
  3. 依內容去重(不同資料夾放了同一張也只算一張)、套用排除清單
  4. 每張:依 EXIF 轉正 → 置中裁成六角比例(640×740)→ 壓縮成 JPG
  5. 依子資料夾輪流排序(輪播才不會連續同一場),輸出 p01.jpg、p02.jpg…
  6. 產 manifest.json,網頁載入時自動讀清單長出輪播

用法(在這個資料夾或任何地方執行都可以):
    python "專案/5_AI 升級計劃/sync_photos.py"            抓新照片 + 重新裁切全部
    python "專案/5_AI 升級計劃/sync_photos.py" --no-download   不連網,只用快取重新裁切

需要一次性安裝:
    pip install gdown pillow
    (若 Drive 裡有 iPhone 的 HEIC 檔,再加 pip install pillow-heif;需要 gdown 6 以上)

跑完後 git commit + push,Cloudflare 會自動上線。
"""
import argparse
import hashlib
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

HERE = Path(__file__).resolve().parent                 # 專案/5_AI 升級計劃
ROOT = HERE.parent.parent                              # 政府補助一站式網站
CACHE = HERE / ".photo_cache"                          # 原檔快取(已加進 .gitignore)
OUT_DIRS = [
    HERE / "2. 產出" / "assets" / "photos",
    ROOT / "course-upgrade" / "assets" / "photos",     # 實際上線的那份
]
EXCLUDE_FILE = HERE / "1. 素材資料" / "照片排除清單.txt"   # 一行一個檔名,# 後面是註解

W, H = 640, 740          # 六角形比例 0.866:1,略留餘裕
QUALITY = 78
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"}
# ────────────────────────────────────────────────────────────────


def log(msg):
    print(msg, flush=True)


def load_excludes():
    names = set()
    if EXCLUDE_FILE.exists():
        for line in EXCLUDE_FILE.read_text(encoding="utf-8").splitlines():
            line = line.split("#", 1)[0].strip()
            if line:
                names.add(line.lower())
    return names


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
    """只下載快取裡還沒有的照片"""
    todo = [f for f in files
            if Path(f.local_path).suffix.lower() in IMG_EXT and not Path(f.local_path).exists()]
    log(f"Drive 共 {len(files)} 個檔案,照片 {sum(1 for f in files if Path(f.local_path).suffix.lower() in IMG_EXT)} 張,"
        f"需要新下載 {len(todo)} 張")
    for i, f in enumerate(todo, 1):
        Path(f.local_path).parent.mkdir(parents=True, exist_ok=True)
        log(f"  [{i}/{len(todo)}] 下載 {Path(f.local_path).relative_to(CACHE)}")
        fetch_file(f.id, Path(f.local_path))


def collect_cached_photos(excludes):
    """從快取收集照片,去重 + 排除,並依子資料夾輪流排序"""
    by_folder = {}
    seen = set()
    skipped_dup = skipped_ex = 0
    for p in sorted(CACHE.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in IMG_EXT:
            continue
        if p.name.lower() in excludes:
            skipped_ex += 1
            continue
        if p.suffix.lower() in {".heic", ".heif"} and not HEIC_OK:
            log(f"  略過 {p.name}(HEIC 需要 pip install pillow-heif)")
            continue
        digest = hashlib.md5(p.read_bytes()).hexdigest()
        if digest in seen:
            skipped_dup += 1
            continue
        seen.add(digest)
        by_folder.setdefault(p.parent, []).append(p)

    # 輪流從每個資料夾各取一張,輪播才不會一整段都是同一場
    ordered = []
    queues = [sorted(v) for _, v in sorted(by_folder.items())]
    while any(queues):
        for q in queues:
            if q:
                ordered.append(q.pop(0))
    log(f"可用照片 {len(ordered)} 張(重複略過 {skipped_dup},排除清單略過 {skipped_ex})")
    return ordered


def crop_hex(src: Path) -> Image.Image:
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    w, h = im.size
    target = W / H
    if w / h > target:
        cw, ch = int(h * target), h
    else:
        cw, ch = w, int(w / target)
    x0, y0 = (w - cw) // 2, (h - ch) // 2
    return im.crop((x0, y0, x0 + cw, y0 + ch)).resize((W, H), Image.LANCZOS)


def write_outputs(photos):
    names = [f"p{i:02d}.jpg" for i in range(1, len(photos) + 1)]
    manifest = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": DRIVE_FOLDER_URL,
        "count": len(names),
        "photos": names,
    }
    rendered = []
    total = 0
    for src, name in zip(photos, names):
        rendered.append((name, crop_hex(src)))
    for out in OUT_DIRS:
        out.mkdir(parents=True, exist_ok=True)
        # 清掉舊的 pNN.jpg,避免留下已刪除的照片
        for old in out.glob("p*.jpg"):
            old.unlink()
        for name, im in rendered:
            path = out / name
            im.save(path, "JPEG", quality=QUALITY, optimize=True, progressive=True)
            total += path.stat().st_size
        (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"寫入 {out}  ({len(names)} 張)")
    log(f"每份約 {total // len(OUT_DIRS) // 1024} KB,manifest.json 已更新")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true", help="不連網,只用快取重新裁切")
    args = ap.parse_args()

    CACHE.mkdir(parents=True, exist_ok=True)
    if not args.no_download:
        download_new(list_drive_files())
    photos = collect_cached_photos(load_excludes())
    if not photos:
        sys.exit("沒有任何照片可用。")
    write_outputs(photos)
    log("完成。接著 git add / commit / push 就會上線。")


if __name__ == "__main__":
    main()
