// 回帰テスト: 品質チェックで見つけた5件の不具合を検証する
// 実行方法: node --test tests/quality_fixes.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');

// モックの状態をテストごとに初期化する
let fetchImpl;
let cacheStore;
let toasts;
let removedBatches;
let logs;
let removeAllError;
let spreadsheet;
let lockHeld;
let fetchAllBatches;
let triggers;
let propertiesStore;

// getRange(row, col, numRows, numCols) と setValues/getValues だけを再現する簡易シート
function makeSheet(name) {
  const grid = [];
  const set = (r, c, v) => {
    while (grid.length < r) grid.push([]);
    grid[r - 1][c - 1] = v;
  };
  const range = (r, c, nr, nc) => ({
    getValues: () => Array.from({ length: nr }, (_, i) =>
      Array.from({ length: nc }, (_, j) => {
        const row = grid[r - 1 + i] || [];
        return row[c - 1 + j] === undefined ? '' : row[c - 1 + j];
      })),
    setValues: (vals) => vals.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))),
    setFormulas: (vals) => vals.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))),
    setValue: (v) => set(r, c, v),
    setFontSize: function () { return this; },
    setFontColor: function () { return this; },
    setHorizontalAlignment: function () { return this; },
    setNumberFormat: function () { return this; },
  });
  return {
    grid,
    getName: () => name,
    getLastRow: () => grid.length,
    clear: () => { grid.length = 0; },
    setFrozenRows: () => {},
    getRange: (r, c, nr = 1, nc = 1) => range(r, c, nr, nc),
  };
}

function makeSpreadsheet() {
  const sheets = new Map();
  return {
    sheets,
    getId: () => 'test-spreadsheet-id',
    toast: (msg, title) => toasts.push({ msg, title }),
    getSheetByName: (name) => sheets.get(name) || null,
    insertSheet: (name) => {
      const sheet = makeSheet(name);
      sheets.set(name, sheet);
      return sheet;
    },
    setActiveSheet: () => {},
  };
}

function resetMocks() {
  fetchImpl = () => { throw new Error('fetch未設定'); };
  cacheStore = new Map();
  toasts = [];
  removedBatches = [];
  logs = [];
  removeAllError = null;
  spreadsheet = makeSpreadsheet();
  lockHeld = false;
  fetchAllBatches = [];
  triggers = [];
  propertiesStore = new Map();
}

function loadModules() {
  resetMocks();

  // Google Apps Script のグローバルをモックする
  // fetchAll は本物と同じく、1件でも例外なら全体が例外になる
  global.UrlFetchApp = {
    fetch: (...args) => fetchImpl(...args),
    fetchAll: (requests) => {
      fetchAllBatches.push(requests.length);
      return requests.map((request) => fetchImpl(request.url, request));
    },
  };
  global.Logger = { log: (msg) => logs.push(String(msg)) };
  const scriptCache = {
    get: (key) => (cacheStore.has(key) ? cacheStore.get(key) : null),
    put: (key, value) => cacheStore.set(key, value),
    getAll: (keys) => {
      const found = {};
      keys.forEach((key) => { if (cacheStore.has(key)) found[key] = cacheStore.get(key); });
      return found;
    },
    putAll: (values) => Object.keys(values).forEach((key) => cacheStore.set(key, values[key])),
    removeAll: (keys) => {
      if (removeAllError) throw removeAllError;
      removedBatches.push([...keys]);
    },
  };
  global.CacheService = {
    getScriptCache: () => scriptCache,
    getUserCache: () => scriptCache,
  };
  global.SpreadsheetApp = {
    getUi: () => ({
      createMenu: () => ({
        addItem: function () { return this; },
        addSeparator: function () { return this; },
        addToUi: () => {},
      }),
    }),
    getActiveSheet: () => ({}),
    getActiveSpreadsheet: () => spreadsheet,
    openById: (id) => (id === spreadsheet.getId() ? spreadsheet : null),
  };
  global.ScriptApp = {
    newTrigger: (handler) => ({
      timeBased: () => ({
        after: (delayMs) => ({
          create: () => {
            const trigger = { getHandlerFunction: () => handler, delayMs };
            triggers.push(trigger);
            return trigger;
          },
        }),
      }),
    }),
    getProjectTriggers: () => [...triggers],
    deleteTrigger: (trigger) => {
      const index = triggers.indexOf(trigger);
      if (index >= 0) triggers.splice(index, 1);
    },
  };
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => (propertiesStore.has(key) ? propertiesStore.get(key) : null),
      setProperty: (key, value) => propertiesStore.set(key, String(value)),
      deleteProperty: (key) => propertiesStore.delete(key),
    }),
  };
  global.LockService = {
    getScriptLock: () => ({
      tryLock: () => {
        if (lockHeld) return false;
        lockHeld = true;
        return true;
      },
      releaseLock: () => { lockHeld = false; },
    }),
  };
  global.Utilities = { formatDate: () => '2024/01/01 00:00:00' };
  global.Parser = {};

  const src = fs.readFileSync(path.join(root, 'finance.gs'), 'utf8')
    + '\n' + fs.readFileSync(path.join(root, 'stock_screener.gs'), 'utf8');
  const factory = new Function(
    src + '; return { get_stock_data, get_stocks_by_page, get_neatprice,'
    + ' getprice_kabutan, clearAllCache, CONFIG, DEFAULT_SETTINGS, getMaxStockPrice,'
    + ' parseDailyVolumes, computeAverageVolume, getAverageVolume, applyVolumeFilter, getCachedStocksByPage,'
    + ' parseAnnualResults, evaluateDividendSafety, getDividendSafety, ensureSettingsSheet, evaluateEarningsQuality,'
    + ' create_pickup_stocks, create_pickup_stocks_from_cache, getProgress, runPickup, TIME_BUDGET_MS, continuePickupFromTrigger,'
    + ' createYahooFinanceFormula: typeof createYahooFinanceFormula === "function"'
    + ' ? createYahooFinanceFormula : undefined };'
  );
  return factory();
}

// UrlFetchApp.fetch の成功レスポンスを作る
function okResponse(html) {
  return { getResponseCode: () => 200, getContentText: () => html };
}

test('終了タグが無い場合は PriceEmpty を返す', () => {
  const m = loadModules();
  fetchImpl = () => okResponse('<span class="kabuka">1,234');
  assert.equal(m.get_neatprice('https://example.com', '<span class="kabuka">'), 'PriceEmpty');
});

test('空ページは株価取得失敗として記録される', () => {
  const m = loadModules();
  fetchImpl = () => okResponse('');
  const result = m.getprice_kabutan('7203');
  assert.equal(result, 'PageEmpty');
  assert.ok(logs.some((line) => line.includes('株価取得失敗')));
});

test('通常の株価は円を取り除いて返す', () => {
  const m = loadModules();
  fetchImpl = () => okResponse('<span class="kabuka">1,234円</span>');
  assert.equal(m.getprice_kabutan('7203'), '1,234');
});

test('キャッシュクリアは取得範囲の全ページを削除する', () => {
  const m = loadModules();
  m.clearAllCache();
  const totalKeys = removedBatches.reduce((sum, batch) => sum + batch.length, 0);
  // 業種33件 × ページ1〜98（pageNo < MAX_PAGES）をすべて消すこと
  assert.equal(totalKeys, 33 * (m.CONFIG.MAX_PAGES - 1));
  assert.ok(removedBatches.every((batch) => batch.length <= 100));
  assert.deepEqual(toasts.map((t) => t.title), ['完了']);
});

test('キャッシュクリア失敗時は成功トーストを表示しない', () => {
  const m = loadModules();
  removeAllError = new Error('削除失敗');
  m.clearAllCache();
  assert.ok(toasts.some((t) => t.title === 'エラー'));
  assert.ok(!toasts.some((t) => t.title === '完了'));
});

test('PER欠損の銘柄はミックス係数で除外される', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/9999/"><abbr title="欠損テスト">欠損</abbr></a></th>`
    + `<td>100</td><td>0</td><td>ー</td><td>1.0倍</td><td>4.0%</td></tr></tbody>`;
  fetchImpl = () => okResponse(row);
  const result = m.get_stocks_by_page('テスト業種', 1, false);
  assert.equal(result.stocks.length, 0);
});

test('PER 1,000倍以上でもカンマで切れずに読み、ミックス係数で除外される', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/3553/"><abbr title="共和レザー">共和レザー</abbr></a></th>`
    + `<td>842</td><td><span class='minus-num'>-20</span><br><span class="text-2xs minus-num">-2.32</span></td>`
    + `<td>1,052.5倍</td><td>0.55倍</td><td>6.18%</td></tr></tbody>`;
  fetchImpl = () => okResponse(row);
  const stock = m.get_stock_data(row.replace(/<\/?tbody>/g, ''));
  assert.equal(stock.PER, 1052.5);
  assert.equal(stock.MIX, 578.875);
  assert.equal(m.get_stocks_by_page('テスト業種', 1, false).stocks.length, 0);
});

test('条件を満たす正常な銘柄は残る', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/7203/"><abbr title="トヨタ自動車">トヨタ</abbr></a></th>`
    + `<td>2,500</td><td><span class='plus-num'>+50</span><br><span class="text-2xs plus-num">2.04</span></td>`
    + `<td>8.0倍</td><td>1.0倍</td><td>4.0%</td></tr></tbody>`;
  fetchImpl = () => okResponse(row);
  const result = m.get_stocks_by_page('テスト業種', 1, false);
  assert.equal(result.stocks.length, 1);
  assert.equal(result.stocks[0].MIX, 8);
});

test('予算上限を超える値がさ株は除外される（株価3,500円×100株=35万円）', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/7203/"><abbr title="値がさテスト">値がさ</abbr></a></th>`
    + `<td>3,500</td><td><span class='plus-num'>+50</span><br><span class="text-2xs plus-num">2.04</span></td>`
    + `<td>8.0倍</td><td>1.0倍</td><td>4.0%</td></tr></tbody>`;
  fetchImpl = () => okResponse(row);
  const result = m.get_stocks_by_page('テスト業種', 1, false, m.DEFAULT_SETTINGS);
  assert.equal(result.stocks.length, 0);
});

test('予算を緩めれば同じ銘柄が通る（settings が即反映される）', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/7203/"><abbr title="値がさテスト">値がさ</abbr></a></th>`
    + `<td>3,500</td><td><span class='plus-num'>+50</span><br><span class="text-2xs plus-num">2.04</span></td>`
    + `<td>8.0倍</td><td>1.0倍</td><td>4.0%</td></tr></tbody>`;
  fetchImpl = () => okResponse(row);
  const settings = Object.assign({}, m.DEFAULT_SETTINGS, { maxBudgetPerStock: 400000 });
  const result = m.get_stocks_by_page('テスト業種', 1, false, settings);
  assert.equal(result.stocks.length, 1);
  assert.equal(m.getMaxStockPrice(m.DEFAULT_SETTINGS), 3000);
});

// 株探の時系列ページ風HTMLを作る（各行: 日付 + 6列 + 売買高）
function kabukaHtml(volumes) {
  const rows = volumes.map((v, i) => {
    const day = String(20 - i).padStart(2, '0');
    return `<tr><th scope="row"><time datetime="2026-11-${day}">26/11/${day}</time></th>`
      + `<td>100</td><td>110</td><td>90</td><td>105</td><td>+5</td><td>5.00</td>`
      + `<td>${v.toLocaleString('en-US')}</td></tr>`;
  });
  return `<html><body><table>${rows.join('')}</table></body></html>`;
}

test('時系列ページから25日平均出来高を計算する', () => {
  const m = loadModules();
  fetchImpl = () => okResponse(kabukaHtml(Array(25).fill(20000)));
  const result = m.getAverageVolume('7203', false);
  assert.equal(result.average, 20000);
  assert.equal(result.daysUsed, 25);
});

test('平均出来高ちょうど下限（10,000株）は除外され、超えれば残る', () => {
  const m = loadModules();
  const stock = { TICKER: '7203' };
  fetchImpl = () => okResponse(kabukaHtml(Array(25).fill(10000)));
  assert.equal(m.applyVolumeFilter([Object.assign({}, stock)], m.DEFAULT_SETTINGS, false).length, 0);
  fetchImpl = () => okResponse(kabukaHtml(Array(25).fill(10001)));
  assert.equal(m.applyVolumeFilter([Object.assign({}, stock)], m.DEFAULT_SETTINGS, false).length, 1);
  fetchImpl = () => okResponse(kabukaHtml(Array(25).fill(9999)));
  assert.equal(m.applyVolumeFilter([Object.assign({}, stock)], m.DEFAULT_SETTINGS, false).length, 0);
});

test('25日に満たない場合は取れた日数で平均し、5日未満は不明として残す', () => {
  const m = loadModules();
  fetchImpl = () => okResponse(kabukaHtml(Array(10).fill(30000)));
  const result = m.getAverageVolume('7203', false);
  assert.equal(result.average, 30000);
  assert.equal(result.daysUsed, 10);
  fetchImpl = () => okResponse(kabukaHtml(Array(3).fill(30000)));
  const few = m.applyVolumeFilter([{ TICKER: '7203' }], m.DEFAULT_SETTINGS, false);
  assert.equal(few.length, 1);
  assert.equal(few[0].AVG_VOLUME, null);
});

test('出来高は10件ずつまとめて取得し、キャッシュ済みの銘柄は取りに行かない', () => {
  const m = loadModules();
  const stocks = Array.from({ length: 23 }, (_, i) => ({ TICKER: String(1000 + i) }));
  const fetched = [];
  fetchImpl = (url) => { fetched.push(url); return okResponse(kabukaHtml(Array(25).fill(20000))); };
  m.applyVolumeFilter(stocks.slice(0, 5).map((s) => Object.assign({}, s)), m.DEFAULT_SETTINGS, true);
  fetched.length = 0;
  fetchAllBatches = [];
  const progress = [];
  const kept = m.applyVolumeFilter(stocks, m.DEFAULT_SETTINGS, true, (done, total) => progress.push(done + '/' + total));
  assert.equal(kept.length, 23);
  assert.equal(fetched.length, 18);
  assert.deepEqual(fetchAllBatches, [5, 10, 3]);
  assert.deepEqual(progress, ['10/23', '20/23', '23/23']);
});

test('出来高の取得に失敗しても銘柄は落とさず不明になる', () => {
  const m = loadModules();
  fetchImpl = () => { throw new Error('通信失敗'); };
  const kept = m.applyVolumeFilter([{ TICKER: '7203' }], m.DEFAULT_SETTINGS, false);
  assert.equal(kept.length, 1);
  assert.equal(kept[0].AVG_VOLUME, null);
});

test('キャッシュ再判定はネット取得せず、新しい閾値だけを適用する', () => {
  const m = loadModules();
  const row = `<tbody><tr><th><a href="/stocks/7203/"><abbr title="トヨタ自動車">トヨタ</abbr></a></th>`
    + `<td>2,500</td><td><span class='plus-num'>+50</span><br><span class="text-2xs plus-num">2.04</span></td>`
    + `<td>8.0倍</td><td>1.0倍</td><td>4.0%</td></tr></tbody>`;
  // 通常取得でキャッシュ（生データ）を溜める
  fetchImpl = () => okResponse(row);
  assert.equal(m.get_stocks_by_page('テスト業種', 1, true, m.DEFAULT_SETTINGS).stocks.length, 1);
  // 以後はネットに出たら失敗させる
  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  // 利回りを 4.5% に上げると、同じ生データでも落ちる
  const strict = Object.assign({}, m.DEFAULT_SETTINGS, { dividendYieldThreshold: 4.5 });
  const filtered = m.getCachedStocksByPage('テスト業種', 1, strict);
  assert.equal(filtered.stocks.length, 0);
  assert.equal(filtered.cacheMiss, false);
  // 緩い閾値なら残る
  assert.equal(m.getCachedStocksByPage('テスト業種', 1, m.DEFAULT_SETTINGS).stocks.length, 1);
});

test('キャッシュが無いページはネットに出ず cacheMiss になる', () => {
  const m = loadModules();
  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  const result = m.getCachedStocksByPage('無い業種', 1, m.DEFAULT_SETTINGS);
  assert.equal(result.cacheMiss, true);
  assert.equal(result.stocks.length, 0);
  // 出来高もキャッシュのみモードでは取得しに行かない
  const vol = m.applyVolumeFilter([{ TICKER: '7203' }], m.DEFAULT_SETTINGS, true, null, true);
  assert.equal(vol.length, 1);
  assert.equal(vol[0].AVG_VOLUME, null);
});

// 株探の財務ページ風HTML（通期業績表）。rows: [period, isForecast, netIncome, eps, dps, ordinaryIncome]
function financeHtml(rows) {
  const trs = rows.map(([period, isForecast, netIncome, eps, dps, ordinaryIncome = '110,000']) =>
    `<tr ><th scope="row" ><span class="kubun1">I&nbsp;&nbsp;${isForecast ? '予&nbsp;' : ' &nbsp;'}</span>${period}&nbsp;&nbsp;</th>`
    + `<td>1,000,000</td><td>100,000</td><td>${ordinaryIncome}</td><td class="high">${netIncome}</td>`
    + `<td>${eps}</td><td>${dps}</td><td class="fb_pdf1"><a>26/05/08</a></td></tr>`);
  return `<table><thead><tr><th scope="col">決算期</th><th>売上高</th><th>営業益</th><th>経常益</th><th>最終益</th>`
    + `<th><span class="help-label" data-help="修正1株益">修正1株益</span></th>`
    + `<th><span class="help-label" data-help="修正1株配">修正1株配</span></th><th>発表日</th></tr></thead>`
    + `<tbody>${trs.join('')}<tr><th scope="row"><span data-help="前期比">前期比</span></th><td>+1</td></tr></tbody></table>`;
}

test('通期業績表から実績と予想を読み取る', () => {
  const m = loadModules();
  const rows = m.parseAnnualResults(financeHtml([
    ['2025.03', false, '4,765,086', '359.6', '90'],
    ['2026.03', false, '3,848,098', '295.3', '95'],
    ['2027.03', true, '3,570,000', '304.3', '100'],
  ]));
  assert.equal(rows.length, 3);
  assert.deepEqual(rows[2], { period: '2027.03', isForecast: true, ordinaryIncome: 110000, netIncome: 3570000, eps: 304.3, dps: 100 });
  assert.equal(rows[0].isForecast, false);
});

test('スマホ版（s.kabutan.jp）の業績表と日足も読める', () => {
  const m = loadModules();
  const yen = (v) => `${v}<span class='pl-px text-10px font-light text-slate-700'>円</span>`;
  const spRow = (period, netIncome, eps, dps) => `<tr\n        >\n          <th class="sticky">\n            ${period}\n            <span class="text-xs"></span>\n          </th>`
    + `<td class="">1,000</td><td class="">100</td><td class="">110</td><td class="">${netIncome}</td><td class="">${yen(eps)}</td>`
    + `<td>${yen(dps)}</td><td><div><a>2026-05-08</a></div></td><td><div class='text-center'>IFRS</div></td></tr>`;
  const finance = `<table><thead><tr><th class="sticky"><span class="text-center">決算期</span></th><th>最終益</th>`
    + `<th><span class="text-center">修正<br>1株配</span></th></tr></thead><tbody>`
    + `<tr><th class="sticky"></th><td colspan="8">過去分を見る</td></tr>`
    + spRow('2026.03', '3,848,098', '295.3', '95') + spRow('予 2027.03', '3,570,000', '304.3', '100')
    + `</tbody></table>`;
  assert.deepEqual(m.parseAnnualResults(finance), [
    { period: '2026.03', isForecast: false, ordinaryIncome: 110, netIncome: 3848098, eps: 295.3, dps: 95 },
    { period: '2027.03', isForecast: true, ordinaryIncome: 110, netIncome: 3570000, eps: 304.3, dps: 100 },
  ]);

  const day = (date, volume) => `<tr>\n<th class="sticky">${date}</th>\n<td>2,930.5</td><td>2,944.5</td><td>2,900.5</td><td>2,900.5</td>`
    + `<td><span class='minus-num'>-30.0</span></td><td><span class='minus-num'>-1.0</span><span>%</span></td>`
    + `<td>${volume}<span class='pl-px text-10px font-light text-slate-700'>株</span></td>\n</tr>`;
  const daily = `<table><tr><th class="px-2">矢印</th><th>意味</th></tr>`
    + day('11/20', '21,035,200') + day('11/19', '19,040,800') + `</table>`;
  assert.deepEqual(m.parseDailyVolumes(daily), [21035200, 19040800]);
});

test('予想配当性向を計算し、上限を超えたら注意に書く', () => {
  const m = loadModules();
  const ok = m.evaluateDividendSafety([
    { period: '2026.03', isForecast: false, netIncome: 1000, eps: 100, dps: 30 },
    { period: '2027.03', isForecast: true, netIncome: 1100, eps: 100, dps: 40 },
  ], m.DEFAULT_SETTINGS);
  assert.equal(ok.payoutRatio, 40);
  assert.deepEqual(ok.flags, []);

  const high = m.evaluateDividendSafety([
    { period: '2026.03', isForecast: false, netIncome: 1000, eps: 100, dps: 80 },
    { period: '2027.03', isForecast: true, netIncome: 1100, eps: 100, dps: 85 },
  ], m.DEFAULT_SETTINGS);
  assert.equal(high.payoutRatio, 85);
  assert.deepEqual(high.flags, ['配当性向 85%']);
});

test('予想赤字・減配予想・過去の減配・減益予想を注意に書く', () => {
  const m = loadModules();
  const result = m.evaluateDividendSafety([
    { period: '2024.03', isForecast: false, netIncome: 1000, eps: 100, dps: 50 },
    { period: '2025.03', isForecast: false, netIncome: 900, eps: 90, dps: 40 },
    { period: '2026.03', isForecast: false, netIncome: 1000, eps: 100, dps: 45 },
    { period: '2027.03', isForecast: true, netIncome: -100, eps: -10, dps: 30 },
  ], m.DEFAULT_SETTINGS);
  assert.equal(result.payoutRatio, null);
  assert.deepEqual(result.flags, ['予想赤字', '減配予想', '2025.03 減配', '減益予想 -110%']);
});

test('財務の取得に失敗しても落とさず注意に書き、キャッシュのみではネットに出ない', () => {
  const m = loadModules();
  fetchImpl = () => { throw new Error('通信失敗'); };
  assert.deepEqual(m.getDividendSafety('7203', m.DEFAULT_SETTINGS, false).flags, ['財務取得失敗（通信エラー）']);
  fetchImpl = () => ({ getResponseCode: () => 403, getContentText: () => '' });
  assert.deepEqual(m.getDividendSafety('7203', m.DEFAULT_SETTINGS, false).flags, ['財務取得失敗（HTTP 403）']);
  fetchImpl = () => okResponse('<html>表のないページ</html>');
  assert.deepEqual(m.getDividendSafety('7203', m.DEFAULT_SETTINGS, false).flags, ['財務取得失敗（表なし）']);
  assert.deepEqual(m.getDividendSafety('7203', m.DEFAULT_SETTINGS, true, true).flags, ['財務データ未取得']);
});

test('財務はキャッシュから settings を変えて再判定できる', () => {
  const m = loadModules();
  fetchImpl = () => okResponse(financeHtml([
    ['2026.03', false, '1,000', '100', '50'],
    ['2027.03', true, '1,100', '100', '60'],
  ]));
  assert.deepEqual(m.getDividendSafety('7203', m.DEFAULT_SETTINGS, true).flags, []);
  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  const strict = Object.assign({}, m.DEFAULT_SETTINGS, { maxPayoutRatio: 50 });
  assert.deepEqual(m.getDividendSafety('7203', strict, true, true).flags, ['配当性向 60%']);
});

test('settings の設定名と説明は定義に同期し、値は残す', () => {
  const m = loadModules();
  const sheet = spreadsheet.insertSheet('settings');
  sheet.getRange(1, 1, 2, 4).setValues([
    ['キー', '設定名', '値', '説明'],
    ['maxPayoutRatio', '【第2段階・予約】配当性向の上限', 60, '未適用'],
  ]);
  m.ensureSettingsSheet(spreadsheet);
  const row = sheet.getRange(2, 1, 1, 4).getValues()[0];
  assert.equal(row[1], '予想配当性向の上限');
  assert.equal(row[2], 60);
  assert.ok(!String(row[3]).includes('未適用'));
});

// 銀行業の1ページ目にだけ1銘柄があり、出来高・財務も取れる状態の fetch
function bankFixtureFetch(url) {
  const bank = encodeURIComponent('銀行業');
  const sectorRow = `<tbody><tr><th><a href="/stocks/8306/"><abbr title="テスト銀行">テスト銀行</abbr></a></th>`
    + `<td>1,500</td><td><span class='plus-num'>+10</span><br><span class="text-2xs plus-num">0.67</span></td>`
    + `<td>8.0倍</td><td>0.8倍</td><td>4.0%</td></tr></tbody>`;
  if (url.includes('/sectors/' + bank + '/?page=1')) return okResponse(sectorRow);
  if (url.includes('/sectors/')) return okResponse('<html></html>');
  if (url.includes('/historical_prices/daily/')) return okResponse(kabukaHtml(Array(25).fill(50000)));
  if (url.includes('/finance/')) {
    return okResponse(financeHtml([
      ['2026.03', false, '1,000', '100', '50'],
      ['2027.03', true, '900', '100', '50'],
    ]));
  }
  throw new Error('想定外のURL: ' + url);
}

test('通常実行は業種・金融の印・注意を出して history に記録し、再判定は記録しない', () => {
  const m = loadModules();
  fetchImpl = bankFixtureFetch;

  const result = m.create_pickup_stocks();
  const summary = '候補 1件（うち注意あり 1件）を pickup に出力しました';
  assert.deepEqual(result, { done: true, message: summary });
  assert.deepEqual(m.getProgress(), { message: summary, percent: 100, done: true });
  // 進捗はサイドバー用のキャッシュに書き、トーストは出さない
  assert.equal(toasts.length, 0);
  const pickup = spreadsheet.getSheetByName('pickup');
  const header = pickup.grid[0];
  const row = pickup.grid[1];
  assert.equal(row[header.indexOf('業種')], '銀行業');
  assert.equal(row[header.indexOf('予想配当性向')], 50);
  const notes = row[header.indexOf('注意')];
  assert.ok(notes.startsWith('金融業'));
  assert.ok(notes.includes('減益予想 -10%'));

  const history = spreadsheet.getSheetByName('history');
  assert.equal(history.getLastRow(), 2);
  assert.equal(history.grid[1][1], '8306');

  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  m.create_pickup_stocks_from_cache();
  assert.equal(history.getLastRow(), 2);
  assert.equal(spreadsheet.getSheetByName('pickup').grid[1][0], '8306');
});

test('キャッシュが空の再判定は終了理由を要約として返し、進捗を完了にする', () => {
  const m = loadModules();
  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  const result = m.create_pickup_stocks_from_cache();
  assert.equal(result.done, true);
  assert.ok(result.message.startsWith('キャッシュにデータがありません'));
  assert.equal(m.getProgress().done, true);
});

test('時間予算を超えたら打ち切って続きを返し、再実行でキャッシュから完了する', () => {
  const m = loadModules();
  let fetchCount = 0;
  fetchImpl = (url) => { fetchCount++; return bankFixtureFetch(url); };

  // 予算を負にすると、最初のページ取得前のチェックで打ち切られる
  const first = m.runPickup(false, -1);
  assert.equal(first.done, false);
  assert.equal(m.getProgress().done, false);
  assert.equal(spreadsheet.getSheetByName('pickup'), null);
  assert.equal(spreadsheet.getSheetByName('history'), null);

  // 2回目は通常の予算で完了する
  const second = m.runPickup(false);
  assert.equal(second.done, true);
  assert.equal(spreadsheet.getSheetByName('history').getLastRow(), 2);

  // 3回目はすべてキャッシュから読むのでネットに出ない
  const before = fetchCount;
  assert.equal(m.runPickup(false).done, true);
  assert.equal(fetchCount, before);
  // 打ち切り・完了のどちらでもロックは解放されている
  assert.equal(lockHeld, false);
});

test('最終益が経常益の90%を超える分だけ PER を補正し、普通の会社と特別損失の会社は変えない', () => {
  const m = loadModules();
  const rows = (ordinaryIncome, netIncome) => [
    { period: '2026.03', isForecast: false, ordinaryIncome: 1000, netIncome: 700, eps: 70, dps: 30 },
    { period: '2027.03', isForecast: true, ordinaryIncome, netIncome, eps: 70, dps: 30 },
  ];
  // 比率0.87（6138 相当）は補正しない
  assert.deepEqual(m.evaluateEarningsQuality(rows(1500, 1300), 3.48, m.DEFAULT_SETTINGS).adjustedPer, 3.48);
  // 比率5.13（544A 相当）は 2.32 × 5.13 ÷ 0.9 ≒ 13.2
  const inflated = m.evaluateEarningsQuality(rows(1410, 7230), 2.32, m.DEFAULT_SETTINGS);
  assert.equal(inflated.adjustedPer, 13.22);
  assert.deepEqual(inflated.flags, ['特別利益でかさ上げ（最終益/経常益 513%）']);
  // 比率0.44（特別損失、7214 相当）は持ち上げない
  assert.equal(m.evaluateEarningsQuality(rows(2250, 1000), 5.19, m.DEFAULT_SETTINGS).adjustedPer, 5.19);
  // 経常赤字予想は判定不能
  const loss = m.evaluateEarningsQuality(rows(-100, 500), 5, m.DEFAULT_SETTINGS);
  assert.equal(loss.adjustedPer, null);
  assert.deepEqual(loss.flags, ['経常赤字予想']);
  // 予想の経常益が無ければ株探PERのまま
  const missing = m.evaluateEarningsQuality(rows(null, 500), 5, m.DEFAULT_SETTINGS);
  assert.equal(missing.adjustedPer, 5);
  assert.deepEqual(missing.flags, ['経常ベース判定不可']);
  // 上限は settings で変えられる
  const loose = Object.assign({}, m.DEFAULT_SETTINGS, { maxNetToOrdinaryRatio: 600 });
  assert.equal(m.evaluateEarningsQuality(rows(1410, 7230), 2.32, loose).adjustedPer, 2.32);
});

// 銀行業の1ページ目に3銘柄。A は比率3（補正MIX 6.67）、B は普通（MIX 5）、C は比率6（補正MIX 13.3）
function qualityFixtureFetch(url) {
  const bank = encodeURIComponent('銀行業');
  const stock = (code, per) => `<tr><th><a href="/stocks/${code}/"><abbr title="銘柄${code}">銘柄${code}</abbr></a></th>`
    + `<td>1,000</td><td><span class='plus-num'>+10</span><br><span class="text-2xs plus-num">1.00</span></td>`
    + `<td>${per}倍</td><td>1.0倍</td><td>4.0%</td></tr>`;
  if (url.includes('/sectors/' + bank + '/?page=1')) {
    return okResponse(`<tbody>${stock('1001', '2.0')}${stock('1002', '5.0')}${stock('1003', '2.0')}</tbody>`);
  }
  if (url.includes('/sectors/')) return okResponse('<html></html>');
  if (url.includes('/historical_prices/daily/')) return okResponse(kabukaHtml(Array(25).fill(50000)));
  const ratioByCode = { 1001: 3, 1002: 0.7, 1003: 6 };
  const code = url.match(/stocks\/(\d+)\/finance/)[1];
  const net = 1000 * ratioByCode[code];
  return okResponse(financeHtml([
    ['2026.03', false, String(net), '100', '30', '1,000'],
    ['2027.03', true, String(net), '100', '30', '1,000'],
  ]));
}

test('補正ミックス係数で閾値を判定し直して並べ、既存 history の見出しに補正PERを足す', () => {
  const m = loadModules();
  fetchImpl = qualityFixtureFetch;
  const oldHistory = spreadsheet.insertSheet('history');
  const oldHeader = ['実行日時','コード','名称','業種','株価','PER','PBR','配当利回','ミックス係数','出来高（25日平均）','予想配当性向','注意','適用設定'];
  oldHistory.getRange(1, 1, 1, oldHeader.length).setValues([oldHeader]);

  m.create_pickup_stocks();
  const pickup = spreadsheet.getSheetByName('pickup');
  const header = pickup.grid[0];
  const codes = pickup.grid.slice(1).map((row) => row[0]);
  assert.deepEqual(codes, ['1002', '1001']);
  const rowA = pickup.grid[2];
  assert.equal(rowA[header.indexOf('PER')], 2);
  assert.equal(rowA[header.indexOf('補正PER')], 6.67);
  assert.equal(rowA[header.indexOf('ミックス係数')], 6.67);
  assert.ok(rowA[header.indexOf('注意')].includes('特別利益でかさ上げ（最終益/経常益 300%）'));

  const history = spreadsheet.getSheetByName('history');
  assert.equal(history.grid[0][13], '補正PER');
  assert.equal(history.getLastRow(), 3);
  assert.equal(history.grid[2][13], 6.67);
});

test('前の実行が動いている間は何もせず busy を返し、history を二重に書かない', () => {
  const m = loadModules();
  fetchImpl = () => { throw new Error('ネット取得禁止'); };
  lockHeld = true;
  const result = m.runPickup(false);
  assert.equal(result.done, false);
  assert.equal(result.busy, true);
  assert.equal(spreadsheet.getSheetByName('history'), null);
  assert.equal(lockHeld, true);
});

test('1回の実行の時間予算は90秒（150秒では区切りの応答が届かなかった）', () => {
  const m = loadModules();
  assert.equal(m.TIME_BUDGET_MS, 90000);
});

test('時間予算で区切ったら継続トリガーを1つだけ予約し、トリガーから完走して予約が消える', () => {
  const m = loadModules();
  fetchImpl = bankFixtureFetch;

  // 予算を負にして即区切る。区切りごとに予約し直してもトリガーは1つだけ
  assert.equal(m.runPickup(false, -1).done, false);
  assert.equal(triggers.length, 1);
  assert.equal(triggers[0].getHandlerFunction(), 'continuePickupFromTrigger');
  assert.equal(m.runPickup(false, -1).done, false);
  assert.equal(triggers.length, 1);
  assert.equal(propertiesStore.get('pickup_spreadsheet_id'), 'test-spreadsheet-id');
  assert.equal(spreadsheet.getSheetByName('pickup'), null);

  // サイドバーを介さずトリガーから続きを実行すると完走し、予約は消える
  m.continuePickupFromTrigger();
  assert.ok(spreadsheet.getSheetByName('pickup'));
  assert.equal(spreadsheet.getSheetByName('history').getLastRow(), 2);
  assert.equal(triggers.length, 0);
  assert.equal(lockHeld, false);
});

test('サイドバーは NetworkError でない失敗でも再接続し、区切りの続きを呼ぶ', () => {
  const html = fs.readFileSync(path.join(root, 'progress.html'), 'utf8');
  const script = html.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace('<?!= JSON.stringify(cacheOnly) ?>', 'false');
  const elements = {};
  const timers = [];
  let calls = 0;
  let onSuccess;
  let onFailure;
  global.document = {
    getElementById: (id) => elements[id]
      || (elements[id] = { style: {}, classList: { add: () => {} }, textContent: '' }),
  };
  global.google = {
    script: {
      host: { close: () => {} },
      run: {
        withSuccessHandler: function (fn) { onSuccess = fn; return this; },
        withFailureHandler: function (fn) { onFailure = fn; return this; },
        runPickup: () => {
          calls++;
          // 1回目: 文面に NetworkError / HTTP 0 を含まない失敗（旧実装はここで止まった）
          if (calls === 1) { onFailure(new Error('接続がタイムアウトしました')); return; }
          // 2回目: 時間予算で区切り、3回目で完了
          if (calls === 2) { onSuccess({ done: false, message: '一旦区切りました' }); return; }
          onSuccess({ done: true, message: '完了しました' });
        },
        getProgress: () => onSuccess(null),
      },
    },
  };
  const originalSetTimeout = global.setTimeout;
  const originalSetInterval = global.setInterval;
  global.setTimeout = (fn) => { timers.push(fn); return timers.length; };
  global.setInterval = () => 0;
  try {
    new Function(script)();
    for (let i = 0; i < 50 && timers.length > 0; i++) timers.shift()();
  } finally {
    global.setTimeout = originalSetTimeout;
    global.setInterval = originalSetInterval;
  }
  assert.equal(calls, 3);
  assert.equal(elements.message.textContent, '完了しました');
  assert.equal(elements.close.style.display, 'inline-block');
});

test('銘柄名のダブルクォートを数式用にエスケープする', () => {
  const m = loadModules();
  assert.equal(typeof m.createYahooFinanceFormula, 'function');
  assert.equal(
    m.createYahooFinanceFormula('7203', 'テスト"株式会社'),
    '=HYPERLINK("https://finance.yahoo.co.jp/quote/7203.T","テスト""株式会社")'
  );
});
