/**
 * 共用 pstrip 輪播 · 從 Drive manifest 建立左右對流 tile 流
 *
 * 用法:在 index.html 任何地方放一個空容器
 *   <div class="pstrip" data-pstrip-auto></div>
 *   <script src="/_shared/pstrip.js" defer></script>
 *
 * 這個 script 會:
 *   1. 讀 /_shared/media_manifest.json
 *   2. 用其中的照片 + 影片 Drive ID,動態組成 2 列 tile
 *   3. 每列的 tile 複製一份達成無縫循環(不改動 HTML 檔)
 *
 * 每個 course 自己定義:
 *   .pstrip { --pw / --ph / --pg / --pr / --pbg / --pb / --ps } CSS 變數
 *   .pstrip-row、.ptile、@keyframes pstrip-l/-r 等既有樣式
 * 這份 script 完全不管視覺 · 只負責填 tile。
 */
(function () {
  'use strict';

  const MANIFEST_URL = '/_shared/media_manifest.json';
  // Drive 直連格式
  const photoUrl = id => `https://drive.google.com/thumbnail?id=${id}&sz=w800`;
  // 影片縮圖(第一幀 · Drive 自動產)
  const videoThumbUrl = id => `https://drive.google.com/thumbnail?id=${id}&sz=w800`;
  // 點了跳到 Drive 新分頁播
  const videoOpenUrl  = id => `https://drive.google.com/file/d/${id}/view`;

  // 每一行要幾個 tile(超過就讓它無限輪播)
  const TILES_PER_ROW = 14;

  function tilePhoto(item) {
    const img = document.createElement('img');
    img.src = photoUrl(item.id);
    img.alt = '課程現場';
    img.loading = 'lazy';
    img.decoding = 'async';
    img.referrerPolicy = 'no-referrer';
    img.addEventListener('error', () => {
      img.parentElement && (img.parentElement.style.display = 'none');
    });
    return img;
  }

  // 影片 = 縮圖 + ▶ 浮標 · 點了跳 Drive 新分頁播
  function tileVideo(item) {
    const wrap = document.createElement('a');
    wrap.href = videoOpenUrl(item.id);
    wrap.target = '_blank';
    wrap.rel = 'noopener';
    wrap.style.cssText = 'display:block;width:100%;height:100%;position:relative;text-decoration:none;';

    const img = document.createElement('img');
    img.src = videoThumbUrl(item.id);
    img.alt = '課程現場影片';
    img.loading = 'lazy';
    img.decoding = 'async';
    img.referrerPolicy = 'no-referrer';
    img.style.cssText = 'width:100%;height:100%;object-fit:cover;display:block;';
    img.addEventListener('error', () => {
      wrap.parentElement && (wrap.parentElement.style.display = 'none');
    });

    const badge = document.createElement('div');
    badge.setAttribute('aria-hidden', 'true');
    badge.style.cssText = [
      'position:absolute',
      'top:50%','left:50%',
      'transform:translate(-50%,-50%)',
      'width:44px','height:44px',
      'border-radius:50%',
      'background:rgba(0,0,0,0.55)',
      'display:flex','align-items:center','justify-content:center',
      'pointer-events:none',
    ].join(';');
    badge.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="#fff" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>';

    wrap.appendChild(img);
    wrap.appendChild(badge);
    return wrap;
  }

  function makeTile(item) {
    const tile = document.createElement('div');
    tile.className = 'ptile';
    tile.style.position = 'relative';   // badge 定位用
    if (item.kind === 'photo') tile.appendChild(tilePhoto(item));
    else tile.appendChild(tileVideo(item));
    return tile;
  }

  // Fisher-Yates shuffle · in-place
  function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }

  // 按 subfolder 分桶 · 桶內洗牌
  function makeBuckets(items) {
    const buckets = {};
    items.forEach(it => {
      const key = it.subfolder || '(root)';
      (buckets[key] = buckets[key] || []).push(it);
    });
    Object.values(buckets).forEach(shuffle);
    return buckets;
  }

  // 一列:從多個 subfolder 桶 round-robin 取 · phaseOffset 讓兩列起點錯開
  // kind='photo' 或 'video' · 加在 item 上
  function roundRobinTake(buckets, want, phaseOffset, kind) {
    const keys = Object.keys(buckets);
    if (!keys.length) return [];
    const out = [];
    let step = 0;
    while (out.length < want) {
      let took = 0;
      for (let k = 0; k < keys.length; k++) {
        const key = keys[(step + phaseOffset + k) % keys.length];
        const b = buckets[key];
        if (b && b.length) {
          out.push({ ...b.shift(), kind });
          took++;
          if (out.length >= want) break;
        }
      }
      if (!took) break;   // 所有桶空了
      step++;
    }
    return out;
  }

  // 把影片均勻散進一列照片中(避免影片全擠一起)
  function sprinkleVideos(photos, videos) {
    if (!videos.length) return photos;
    const total = photos.length + videos.length;
    const positions = new Set();
    // 想要的 video 位置 · 大致等距分布
    for (let i = 0; i < videos.length; i++) {
      const pos = Math.floor(((i + 0.5) * total) / videos.length);
      positions.add(Math.min(total - 1, Math.max(0, pos)));
    }
    const out = [];
    let pi = 0, vi = 0;
    for (let i = 0; i < total; i++) {
      if (positions.has(i) && vi < videos.length) out.push(videos[vi++]);
      else if (pi < photos.length) out.push(photos[pi++]);
      else if (vi < videos.length) out.push(videos[vi++]);
    }
    return out;
  }

  function fillPstrip(container, manifest) {
    const allPhotos = manifest.photos || [];
    const allVideos = manifest.videos || [];
    if (!allPhotos.length && !allVideos.length) {
      container.style.display = 'none';
      return;
    }

    // 兩列各自 round-robin · 起點錯開 1 格(避免左右平行同 subfolder)
    const photoBuckets = makeBuckets(allPhotos);
    const videoBuckets = makeBuckets(allVideos);

    // 每列的 photo / video 配額
    const videoPerRow = Math.min(
      Math.ceil(TILES_PER_ROW / 4),             // 大概 4 格 1 支
      Math.ceil(allVideos.length / 2)           // 影片總數也要撐得起兩列
    );
    const photoPerRow = TILES_PER_ROW - videoPerRow;

    const leftPhotos  = roundRobinTake(photoBuckets, photoPerRow, 0, 'photo');
    const leftVideos  = roundRobinTake(videoBuckets, videoPerRow, 0, 'video');
    const rightPhotos = roundRobinTake(photoBuckets, photoPerRow, 1, 'photo');
    const rightVideos = roundRobinTake(videoBuckets, videoPerRow, 1, 'video');

    const leftItems  = sprinkleVideos(leftPhotos,  leftVideos);
    const rightItems = sprinkleVideos(rightPhotos, rightVideos);

    container.innerHTML = '';

    function buildRow(items, direction) {
      const row = document.createElement('div');
      row.className = 'pstrip-row pstrip-' + direction;
      // 兩份達成無縫循環
      [...items, ...items].forEach((it, idx) => {
        const tile = makeTile(it);
        if (idx >= items.length) tile.setAttribute('aria-hidden', 'true');
        row.appendChild(tile);
      });
      return row;
    }

    container.appendChild(buildRow(leftItems, 'left'));
    container.appendChild(buildRow(rightItems, 'right'));
  }

  async function init() {
    const containers = document.querySelectorAll('[data-pstrip-auto]');
    if (!containers.length) return;

    try {
      const res = await fetch(MANIFEST_URL, { cache: 'no-cache' });
      if (!res.ok) throw new Error('fetch failed: ' + res.status);
      const manifest = await res.json();
      containers.forEach(c => fillPstrip(c, manifest));
    } catch (err) {
      console.warn('[pstrip] 讀 media_manifest.json 失敗:', err);
      containers.forEach(c => c.style.display = 'none');
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
