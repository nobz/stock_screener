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

function resetMocks() {
  fetchImpl = () => { throw new Error('fetch未設定'); };
  cacheStore = new Map();
  toasts = [];
  removedBatches = [];
  logs = [];
  removeAllError = null;
}

function loadModules() {
  resetMocks();

  // Google Apps Script のグローバルをモックする
  global.UrlFetchApp = { fetch: (...args) => fetchImpl(...args) };
  global.Logger = { log: (msg) => logs.push(String(msg)) };
  global.CacheService = {
    getScriptCache: () => ({
      get: (key) => (cacheStore.has(key) ? cacheStore.get(key) : null),
      put: (key, value) => cacheStore.set(key, value),
      removeAll: (keys) => {
        if (removeAllError) throw removeAllError;
        removedBatches.push([...keys]);
      },
    }),
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
    getActiveSpreadsheet: () => ({ toast: (msg, title) => toasts.push({ msg, title }) }),
  };
  global.Utilities = { formatDate: () => '2024/01/01 00:00:00' };
  global.Parser = {};

  const src = fs.readFileSync(path.join(root, 'finance.gs'), 'utf8')
    + '\n' + fs.readFileSync(path.join(root, 'stock_screener.gs'), 'utf8');
  const factory = new Function(
    src + '; return { get_stock_data, get_stocks_by_page, get_neatprice,'
    + ' getprice_kabutan, clearAllCache, CONFIG,'
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

test('銘柄名のダブルクォートを数式用にエスケープする', () => {
  const m = loadModules();
  assert.equal(typeof m.createYahooFinanceFormula, 'function');
  assert.equal(
    m.createYahooFinanceFormula('7203', 'テスト"株式会社'),
    '=HYPERLINK("https://finance.yahoo.co.jp/quote/7203.T","テスト""株式会社")'
  );
});
