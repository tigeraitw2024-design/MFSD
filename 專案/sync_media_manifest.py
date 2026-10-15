# -*- coding: utf-8 -*-
"""
sync_media_manifest.py — 讀 Google Drive 公開資料夾 · 產 media_manifest.json

跟舊的 sync_photos.py 不同:
  * 完全不下載檔案(省 Cloudflare 空間)
  * 只產一份 JSON 清單 · 存檔案 ID + 類型 + 原檔名
  * 前端 pstrip.js 讀這份清單 · 直接 embed Drive URL

用法:
    python "專案/sync_media_manifest.py"

需要:
    pip install gdown requests

跑完後 git commit + push · 7 + 新增課程網頁自動載到最新 media。
"""
import json
import sys
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

try:
    import requests
    from gdown.download_folder import _download_and_parse_google_drive_link, _parse_google_drive_folder_id
except ImportError:
    sys.exit('缺少套件 · 請執行:pip install gdown requests')

# ───────────────────────────── 設定 ─────────────────────────────
DRIVE_FOLDER_URL = 'https://drive.google.com/drive/folders/1dIqO4XVL2hIDVoPgQT46q-Tnj3tdBraf'

HERE = Path(__file__).resolve().parent        # 專案/
ROOT = HERE.parent                            # 政府補助一站式網站/
OUT_DIR = ROOT / '_shared'
OUT_FILE = OUT_DIR / 'media_manifest.json'

PHOTO_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.heic'}
VIDEO_EXTS = {'.mov', '.mp4', '.webm', '.m4v'}


def walk(node, parent='', bucket_photos=None, bucket_videos=None, bucket_subfolders=None):
    """遞迴走訪 Drive 樹 · 收集所有 photo/video 檔案 ID。"""
    is_folder = (node.type == 'application/vnd.google-apps.folder')
    this_name = node.name
    this_path = (parent + '/' + this_name) if parent else this_name
    if is_folder:
        if parent:   # 子資料夾才記錄
            bucket_subfolders.add(this_path)
        for child in getattr(node, 'children', []) or []:
            walk(child, this_path, bucket_photos, bucket_videos, bucket_subfolders)
        return
    # 檔案 · 看副檔名分類
    ext = ''
    if '.' in this_name:
        ext = '.' + this_name.rsplit('.', 1)[1].lower()
    entry = {
        'id': node.id,
        'name': this_name,
        'subfolder': parent,   # 子資料夾路徑(排序用)
    }
    if ext in PHOTO_EXTS:
        bucket_photos.append(entry)
    elif ext in VIDEO_EXTS:
        bucket_videos.append(entry)


def interleave_by_subfolder(items):
    """依子資料夾輪流排序 · 讓 pstrip 不會連續同一場。"""
    buckets = {}
    for it in items:
        buckets.setdefault(it['subfolder'], []).append(it)
    # Round-robin
    out = []
    while any(buckets.values()):
        for key in list(buckets.keys()):
            if buckets[key]:
                out.append(buckets[key].pop(0))
    return out


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f'📂 讀 Drive 資料夾:{DRIVE_FOLDER_URL}')
    folder_id = _parse_google_drive_folder_id(url=DRIVE_FOLDER_URL)
    sess = requests.Session()
    root = _download_and_parse_google_drive_link(sess=sess, folder_id=folder_id, quiet=True, verify=True)

    photos, videos = [], []
    subfolders = set()
    walk(root, '', photos, videos, subfolders)

    print(f'📸 照片 {len(photos)} 張 · 🎥 影片 {len(videos)} 支 · 📁 子資料夾 {len(subfolders)}')

    photos_ordered = interleave_by_subfolder(photos)
    videos_ordered = interleave_by_subfolder(videos)

    manifest = {
        'generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'source_url': DRIVE_FOLDER_URL,
        'root_name': root.name,
        'subfolders': sorted(subfolders),
        'counts': {'photos': len(photos), 'videos': len(videos)},
        'photos': photos_ordered,
        'videos': videos_ordered,
    }

    OUT_FILE.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'✅ 寫入 {OUT_FILE}')
    print(f'   檔案大小:{OUT_FILE.stat().st_size / 1024:.1f} KB')
    print()
    print('下一步:git add _shared/ && git commit && git push')
    print('前端 pstrip.js 會自動載這份 manifest')


if __name__ == '__main__':
    main()
