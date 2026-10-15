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
  // iframe preview · 保證能播 · 會有 Drive logo
  const videoIframeUrl = id => `https://drive.google.com/file/d/${id}/preview`;

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

  // 影片 iframe · 保證能播 · 會顯示 Drive 的控制列
  function tileVideo(item) {
    const iframe = document.createElement('iframe');
    iframe.src = videoIframeUrl(item.id);
    iframe.allow = 'autoplay';
    iframe.loading = 'lazy';
    iframe.setAttribute('allowfullscreen', '');
    iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;';
    return iframe;
  }

  function makeTile(item) {
    const tile = document.createElement('div');
    tile.className = 'ptile';
    tile.style.position = 'relative';   // badge 定位用
    if (item.kind === 'photo') tile.appendChild(tilePhoto(item));
    else tile.appendChild(tileVideo(item));
    return tile;
  }

  // 把 photos 跟 videos 交錯合併(大概每 4 張照片插 1 支影片)
  function mergeMedia(photos, videos) {
    const merged = [];
    let vi = 0;
    const videoEvery = Math.max(1, Math.ceil(photos.length / Math.max(1, videos.length)));
    for (let i = 0; i < photos.length; i++) {
      merged.push({ ...photos[i], kind: 'photo' });
      if (vi < videos.length && (i + 1) % videoEvery === 0) {
        merged.push({ ...videos[vi++], kind: 'video' });
      }
    }
    // 剩下的影片附在尾巴
    while (vi < videos.length) merged.push({ ...videos[vi++], kind: 'video' });
    return merged;
  }

  function fillPstrip(container, manifest) {
    const merged = mergeMedia(manifest.photos || [], manifest.videos || []);
    if (!merged.length) {
      container.style.display = 'none';
      return;
    }

    // 分成兩列:奇數 index 左列、偶數 index 右列
    const leftItems = merged.filter((_, i) => i % 2 === 0).slice(0, TILES_PER_ROW);
    const rightItems = merged.filter((_, i) => i % 2 === 1).slice(0, TILES_PER_ROW);

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
