// ============================================================
//  SyncEngine v1 — 前端差量同步模組 (join-up with api_sync.gs v4)
//  能力：版本差量拉取 / 行級 ETag / checksum 對帳 / 寫入衝突自動 rebase
//  ★ 整合方式（見 README）：index.html 內 <script src=".../sync.js"></script>
//    並把 loadData() 的全量刷新改為 Sync.refresh()
// ============================================================

var Sync = (function () {

  // ---------- 配置（沿用 index.html 既有常數）----------
  var API_URL = typeof API_URL_ !== 'undefined' ? API_URL_ :
    'https://script.google.com/macros/library/d/1IcjDkLBSm-cXeE8skbjeVYRuBJle6nGF4GBP5mnz2VDOKoAhGJVO_MyA/19';
  var CDN_URL = typeof CDN_URL_ !== 'undefined' ? CDN_URL_ :
    'https://raw.githubusercontent.com/charlizelai/product-search/main/data/products.json';
  var LS_VERSION_KEY = 'sync_version';     // 本地已知的全域版本號
  var LS_DATA_KEY = 'products_v2';         // 沿用現有本地快取
  var LS_SNAPSHOT_VER_KEY = 'sync_snapshot_ver'; // 本地快照對應的 server version
  var MAX_RETRIES = 3;
  var MAX_RESYNC_DATA_AGE = 3 * 24 * 60 * 60 * 1000; // 3 天，過期就走 CODE 端點

  // ---------- 狀態 ----------
  var localVersion = 0;
  var snapshotVer = 0;
  var snapshotTime = 0;
  var resyncing = false;

  function safeLS(key, val) {
    try {
      if (val === undefined) return localStorage.getItem(key);
      localStorage.setItem(key, val);
    } catch (e) { return null; }
  }

  function nowVersion() { return localVersion; }

  // ---------- 啟動：從本地快取及時渲染，再背景差量刷新 ----------
  function boot() {
    try {
      localVersion = Number(safeLS(LS_VERSION_KEY)) || 0;
      snapshotVer = Number(safeLS(LS_SNAPSHOT_VER_KEY)) || 0;
      snapshotTime = Number(safeLS('sync_snapshot_time')) || 0;
      var cached = safeLS(LS_DATA_KEY);
      if (cached) {
        var data = JSON.parse(cached);
        if (Array.isArray(data) && window.__onDataLoaded) window.__onDataLoaded(data);
      }
    } catch (e) { console.warn('Sync.boot cache error', e); }
    refresh();
    // 每 5 分鐘背景差量檢查（低頻、輕量）
    setInterval(refresh, 5 * 60 * 1000);
  }

  // ---------- 核心刷新：差量 > 全量快照 > CDN ----------
  function refresh() {
    var chain = Promise.resolve();
    // 快照過舊 => 直接走全量修復，避免基於太舊版本 diff 失敗
    var needsFull = !snapshotTime || (Date.now() - snapshotTime > MAX_RESYNC_DATA_AGE);
    if (needsFull && snapshotVer === localVersion) {
      chain = chain.then(fullResync).catch(function () { return loadCdnFallback(); });
    } else {
      chain = chain
        .then(function () { return deltaFetch(localVersion); })
        .then(function (d) { applyDelta(d); })
        .catch(function (e) { console.warn('Sync delta failed, resyncing…', e); return fullResync(); });
    }
    return chain;
  }

  // ---------- 差量拉取 ----------
  function deltaFetch(since) {
    var token = (typeof TOKEN !== 'undefined' ? TOKEN : '');
    var url = API_URL + '?action=sync&since=' + encodeURIComponent(since) +
              '&token=' + encodeURIComponent(token);
    return fetch(url).then(function (r) { return r.json(); }).then(function (d) {
      if (d.error) throw new Error(d.error);
      return d;
    });
  }

  // ---------- 差量合併 + checksum 對帳 ----------
  function applyDelta(d) {
    if (!d || d.version === undefined) throw new Error('bad_delta');
    if (d.since !== localVersion && d.changed.length !== 0 && resyncing !== false) {
      // 版本錯位（例如另一台裝置已推進版本）→ 走全量
      return fullResync();
    }

    var data = getData();
    var byId = indexById(data);

    (d.changed || []).forEach(function (upd) {
      byId[upd.id] = normalizeProduct(upd);
    });
    (d.deleted || []).forEach(function (id) { delete byId[id]; });

    var merged = Object.keys(byId).map(function (k) { return byId[k]; });

    // checksum 對帳：不匹配 → 不信任 diff，觸發一次全量修復
    if (!verifyChecksum(d, merged)) {
      console.warn('⚠️ checksum mismatch → full resync');
      return fullResync();
    }

    localVersion = d.version;
    safeLS(LS_DATA_KEY, JSON.stringify(merged));
    safeLS(LS_VERSION_KEY, String(localVersion));
    if (window.__onDataLoaded) window.__onDataLoaded(merged);
  }

  function verifyChecksum(d, merged) {
    // 從 diff 內容重建同 hashing 順序，接近就算 pass；
    // 如後端有提供全量 hash 陣列，可用 merged 對齊
    var notChecked = (d.checksumChunks || null);
    if (!notChecked) return true; // 後端暫未送 checksumChunks 就先不做嚴格對帳
    var joined = merged.map(function (p) { return p.id + ':' + (p.etag || ''); });
    return joined.sort().join('|') === notChecked.sort().join('|');
  }

  // ---------- 全量快照修復（userId-level strict, 每次只跑一個）----------
  function fullResync() {
    if (resyncing) return Promise.resolve();
    resyncing = true;
    return fetch(API_URL + '?action=list&token=' +
        encodeURIComponent(typeof TOKEN !== 'undefined' ? TOKEN : ''))
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.error || !Array.isArray(d.data)) throw new Error('list_error');
        localVersion = d.version || 0;
        snapshotVer = localVersion;
        snapshotTime = Date.now();
        safeLS(LS_DATA_KEY, JSON.stringify(d.data));
        safeLS(LS_VERSION_KEY, String(localVersion));
        safeLS(LS_SNAPSHOT_VER_KEY, String(snapshotVer));
        safeLS('sync_snapshot_time', String(snapshotTime));
        if (window.__onDataLoaded) window.__onDataLoaded(d.data);
        console.log('✅ SyncEngine full resync done, v=' + localVersion);
      })
      .finally(function () { resyncing = false; });
  }

  // ---------- CDN fallback（離線可用、無 token 也能拉靜態快照）----------
  function loadCdnFallback() {
    return fetch(CDN_URL).then(function (r) { return r.json(); }).then(function (list) {
      if (Array.isArray(list) && window.__onDataLoaded) window.__onDataLoaded(list);
      console.log('📦 SyncEngine loaded CDN snapshot');
    });
  }

  // ---------- 寫入：through WriteQueue + 衝突自動 rebase ----------
  function save(product, extraSource) {
    var writes = [{
      id: product.id,
      etag: product.etag || '',      // ★ 編輯前看到的指紋（從 cache 中還原）
      fields: buildFieldPatch(product),
      source: extraSource || 'ui'
    }];
    return postQueue(writes).then(function (res) {
      if (!res.ok) throw new Error(res.error || 'write_failed');
      var first = (res.results || [])[0];
      if (first && first.ok) {
        localVersion = res.version;
        safeLS(LS_VERSION_KEY, String(localVersion));
        // 樂觀更新本地快取
        optimisticUpdate(first.id, product);
        return first;
      }
      if (first && first.reason === 'stale') {
        // ★ 伺服器回傳最新行，自動成本零 rebase
        optimisticUpdate(first.id, first.current);
        throw { stale: true, current: first.current,
                message: '⚠️ 此筆已被其他人更新，已同步最新版本' };
      }
      throw new Error(first && first.reason ? first.reason : 'unknown');
    });
  }

  function postQueue(writes, attempt) {
    attempt = attempt || 0;
    return fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'syncWrite', token: (typeof TOKEN !== 'undefined' ? TOKEN : ''), writes: writes })
    }).then(function (r) { return r.json(); })
      .catch(function () {
        if (attempt < MAX_RETRIES) {
          var wait = 400 * Math.pow(2, attempt); // 400/800/1600ms 指數退避
          return new Promise(function (res) { setTimeout(res, wait); })
            .then(function () { return postQueue(writes, attempt + 1); });
        }
        return { ok: false, error: 'network_after_retries' };
      });
  }

  // ---------- 工具 ----------
  function indexById(list) {
    var m = {}; (list || []).forEach(function (p) { m[p.id] = p; }); return m;
  }
  function normalizeProduct(raw) {
    var FIELDS = typeof FIELDS !== 'undefined' ? FIELDS :
      ['id','barcode','name','name_zh','price','original_price','image','category_en',
       'category_zh','description','temp_location','combine','warehouse','island',
       'aisle','crosslane','shelfnumber','floor'];
    var out = {};
    FIELDS.forEach(function (f) { out[f] = raw[f] !== undefined ? raw[f] : ''; });
    out._last_modified = raw._last_modified || '';
    out._last_source = raw._last_source || '';
    out.etag = raw.etag || '';
    out.locked = raw.locked || false;
    return out;
  }
  function buildFieldPatch(p) {
    var patch = {};
    Object.keys(p).forEach(function (k) {
      if (k === 'etag' || k === 'version' || k === '_last_modified' || k === '_last_source') return;
      patch[k] = p[k];
    });
    return patch;
  }
  function optimisticUpdate(id, product) {
    var data = getData();
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i].id) === String(id)) { data[i] = product; found = true; break; }
    }
    if (!found) data.push(product);
    safeLS(LS_DATA_KEY, JSON.stringify(data));
    if (window.__onDataLoaded) window.__onDataLoaded(data);
  }
  function getData() {
    try { return JSON.parse(safeLS(LS_DATA_KEY)) || []; }
    catch (e) { return []; }
  }

  // ---------- Public API ----------
  return {
    boot: boot, refresh: refresh, nowVersion: nowVersion,
    fullResync: fullResync, save: save
  };
})();

// ★ 開機自動啟動（ 在 index.html 底端 app 初始化後一行：  Sync.boot();  ）
