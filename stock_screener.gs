/**
 * 設定定数（閾値の正は settings シート。ここはシートが無い／壊れたときの既定値）
 */
const CONFIG = {
  MAX_PAGES: 99,
  GRAHAM_MIX_THRESHOLD: 11,
  DIVIDEND_YIELD_THRESHOLD: 3.5,
  CACHE_INDUSTRY_DURATION: 21600, // 6時間
  CACHE_STOCK_DURATION: 3600, // 1時間
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
  CACHE_CLEAR_MAX_INDUSTRY: 50,
  CACHE_BATCH_SIZE: 100
};

const SETTINGS_SHEET_NAME = 'settings';
const OUTPUT_SHEET_NAME = 'pickup';

/**
 * settings シートの定義。キーはコード内での名前、値は C 列でユーザーが編集する。
 * stage2 系は第2段階スクリーニング（個別財務ページ取得）導入時に使う予約項目で、
 * 現時点では判定に使わない。シート上の説明にもその旨を書く。
 */
const SETTINGS_DEFINITIONS = [
  { key: 'maxBudgetPerStock', label: '1銘柄の予算上限', value: 300000, description: '円。株価×単元株数がこれを超える銘柄（値がさ株）を除外する' },
  { key: 'unitShares', label: '単元株数', value: 100, description: '株。通常は100。予算判定と購入金額の計算に使う' },
  { key: 'grahamMixThreshold', label: 'グレアム・ミックス係数の上限', value: CONFIG.GRAHAM_MIX_THRESHOLD, description: 'PER×PBR がこれ未満を割安とする（本家グレアム基準は22.5、現行は厳しめの11）' },
  { key: 'dividendYieldThreshold', label: '配当利回りの下限', value: CONFIG.DIVIDEND_YIELD_THRESHOLD, description: '％。これ以上を対象とする' },
  { key: 'minAverageVolume', label: '25日平均出来高の下限', value: 10000, description: '株。直近25営業日の平均出来高がこれ以下の銘柄を除外する（流動性フィルタ）' },
  { key: 'minRoe', label: '【第2段階・予約】ROE の下限', value: 8, description: '％。今の稼ぐ力の判定用。個別財務取得を入れるまでは未適用' },
  { key: 'minOperatingMargin', label: '【第2段階・予約】営業利益率の下限', value: 5, description: '％。業種差が大きいため最終的には業種内比較にする想定。未適用' },
  { key: 'minEquityRatio', label: '【第2段階・予約】自己資本比率の下限', value: 40, description: '％。金融業など業種によっては例外扱いにする想定。未適用' },
  { key: 'maxPayoutRatio', label: '【第2段階・予約】配当性向の上限', value: 70, description: '％。高利回りでも稼ぎ以上に配っていないかの判定用。未適用' },
];

const DEFAULT_SETTINGS = SETTINGS_DEFINITIONS.reduce((acc, def) => {
  acc[def.key] = def.value;
  return acc;
}, {});

/**
 * settings シートが無ければ作り、不足している項目行だけを末尾に足す。
 * 既存の値（C列）は絶対に上書きしない。
 */
function ensureSettingsSheet(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['キー', '設定名', '値', '説明']]);
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const knownKeys = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .forEach((row) => { if (row[0]) knownKeys.add(String(row[0])); });
  }

  const missing = SETTINGS_DEFINITIONS.filter((def) => !knownKeys.has(def.key));
  if (missing.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, missing.length, 4).setValues(
      missing.map((def) => [def.key, def.label, def.value, def.description])
    );
  }
  return sheet;
}

/**
 * settings シートから設定を読む。読めない／0以下・数値でない値は既定値に倒す
 * （予算が NaN になると全銘柄が黙って除外される事故を防ぐため）。
 */
function getSettings(ss) {
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  try {
    const sheet = ensureSettingsSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return settings;

    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (const def of SETTINGS_DEFINITIONS) {
      const row = rows.find((r) => String(r[0]) === def.key);
      if (!row) continue;
      const value = Number(row[2]);
      if (Number.isFinite(value) && value > 0) {
        settings[def.key] = value;
      } else {
        Logger.log('settings の値が不正なため既定値を使います: ' + def.key + ' = ' + row[2]);
      }
    }
  } catch (e) {
    Logger.log('settings 読み込みエラー（既定値を使用）: ' + e.toString());
  }
  settings.unitShares = Math.max(1, Math.round(settings.unitShares));
  return settings;
}

/** 予算上限から逆算した株価の上限（円） */
function getMaxStockPrice(settings) {
  return Math.floor(settings.maxBudgetPerStock / settings.unitShares);
}

/**
 * 第1段階の合否判定。PER/PBR 欠損（MIX=null）は判定不能として除外する。
 */
function passesScreening(stock, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const price = Number(stock.PRICE);
  return stock.MIX !== null
    && stock.MIX < s.grahamMixThreshold
    && stock.DPS >= s.dividendYieldThreshold
    && Number.isFinite(price) && price > 0
    && price * s.unitShares <= s.maxBudgetPerStock;
}

function applyScreening(stocks, settings) {
  return (stocks || []).filter((stock) => passesScreening(stock, settings));
}

/** 出来高の平均を取る営業日数と、平均として認める最低日数 */
const VOLUME_AVERAGE_DAYS = 25;
const VOLUME_MIN_DAYS = 5;

/**
 * 時系列ページのHTMLから日足の売買高を新しい順に抽出する。
 * 株探の日足行: <tr><th scope="row"><time datetime="...">YY/MM/DD</time></th>
 *   <td>始値</td>...<td>売買高</td></tr>（最終 td が売買高）
 */
function parseDailyVolumes(html) {
  const volumes = [];
  const trMatches = String(html || '').match(/<tr>[\s\S]*?<\/tr>/g) || [];
  for (const row of trMatches) {
    if (!/<time\s+datetime=/.test(row)) continue;
    const tdMatches = row.match(/<td>([\s\S]*?)<\/td>/g);
    if (!tdMatches || tdMatches.length < 7) continue;
    const volumeText = removeHtmlTags(tdMatches[tdMatches.length - 1]).replace(/,/g, '').trim();
    if (!/^\d+$/.test(volumeText)) continue;
    volumes.push(Number(volumeText));
  }
  return volumes;
}

/** 売買高リストの先頭 days 件の平均。最低日数に満たなければ null（不明） */
function computeAverageVolume(volumes, days) {
  const target = (volumes || []).slice(0, days || VOLUME_AVERAGE_DAYS);
  if (target.length < VOLUME_MIN_DAYS) return null;
  const sum = target.reduce((acc, v) => acc + v, 0);
  return Math.round(sum / target.length);
}

/**
 * 銘柄の直近25営業日の平均出来高を取得する。戻り値: { average, daysUsed }
 * 取得失敗・日数不足のときは average: null（不明）。
 * キャッシュは銘柄単位で1時間。clearAllCache の対象外（GASでは全キー列挙が
 * できないため）。実害は最大1時間古い出来高で判定すること。
 */
function getAverageVolume(ticker, useCache = true, cacheOnly = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('volume', ticker);

  if (useCache || cacheOnly) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      return JSON.parse(cached);
    }
  }
  // キャッシュのみモードではネット取得しない（出来高は不明として扱う）
  if (cacheOnly) {
    return { average: null, daysUsed: 0 };
  }

  const url = 'https://kabutan.jp/stock/kabuka?code=' + encodeURIComponent(ticker);
  try {
    const volumes = parseDailyVolumes(fetchWithUserAgent(url));
    const average = computeAverageVolume(volumes, VOLUME_AVERAGE_DAYS);
    const result = { average: average, daysUsed: Math.min(volumes.length, VOLUME_AVERAGE_DAYS) };
    if (average !== null) {
      cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_STOCK_DURATION);
    }
    return result;
  } catch (e) {
    Logger.log('出来高取得エラー (' + ticker + '): ' + e.toString());
    return { average: null, daysUsed: 0 };
  }
}

/**
 * 第1.5段階: 平均出来高による流動性フィルタ。
 * 平均が settings.minAverageVolume 以下なら除外する。取得できなかった銘柄は
 * 落とさず、AVG_VOLUME=null のまま残す（株探側の変更で全滅する事故を防ぐ）。
 */
function applyVolumeFilter(stocks, settings, useCache = true, onProgress, cacheOnly = false) {
  const s = settings || DEFAULT_SETTINGS;
  const kept = [];
  const dropped = [];
  (stocks || []).forEach((stock, index) => {
    const result = getAverageVolume(stock.TICKER, useCache, cacheOnly);
    stock.AVG_VOLUME = result.average;
    if (result.average !== null && result.average <= s.minAverageVolume) {
      dropped.push(stock.TICKER);
    } else {
      kept.push(stock);
    }
    if (onProgress) onProgress(index + 1, stocks.length);
  });
  if (dropped.length > 0) {
    Logger.log('出来高フィルタで除外: ' + dropped.length + '件 (' + dropped.join(', ') + ')');
  }
  return kept;
}

/**
 * キャッシュキーを生成
 */
function getCacheKey(type, ...params) {
  return type + '_' + params.join('_');
}

/**
 * User-Agentを付けてHTTPリクエストを実行
 */
function fetchWithUserAgent(url) {
  const options = {
    'headers': {
      'User-Agent': CONFIG.USER_AGENT
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error('HTTP ' + responseCode + ' for URL: ' + url);
    }
    
    return response.getContentText("utf-8");
  } catch(e) {
    Logger.log('HTTPリクエストエラー: ' + url + ' - ' + e.toString());
    throw e;
  }
}

/**
 * HTMLタグを除去して改行も削除
 */
function removeHtmlTags(text) {
  return text.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g, '').replace(/\r\n/g, '');
}

/**
 * スプレッドシート起動時にカスタムメニューを追加
 * メニューから銘柄データ更新とキャッシュクリアを実行可能
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('株価取得')
    .addItem('銘柄データ更新', 'create_pickup_stocks')
    .addItem('キャッシュで再判定（取得しない）', 'create_pickup_stocks_from_cache')
    .addItem('settingsシートを開く', 'openSettingsSheet')
    .addSeparator()
    .addItem('キャッシュクリア', 'clearAllCache')
    .addToUi();
}

/**
 * キャッシュを完全にクリアする関数
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    const iList = create_industrycodelist(false);
    let keys = [];

    // 取得範囲（1ページ〜MAX_PAGES未満）と揃えて全ページ分を削除する
    for (const industry of iList) {
      for (let p = 1; p < CONFIG.MAX_PAGES; p++) {
        keys.push(getCacheKey('stocks', encodeURIComponent(industry.CODE), p));
      }
    }

    // 100個ずつ削除（removeAllの制限対策）
    for (let i = 0; i < keys.length; i += CONFIG.CACHE_BATCH_SIZE) {
      const batch = keys.slice(i, i + CONFIG.CACHE_BATCH_SIZE);
      cache.removeAll(batch);
    }

    Logger.log('キャッシュを完全にクリアしました: ' + keys.length + '個のキー');
    SpreadsheetApp.getActiveSpreadsheet().toast('キャッシュをクリアしました', '完了', 3);
  } catch(e) {
    const errorMsg = 'キャッシュクリアエラー: ' + e.toString();
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, 'エラー', 10);
  }
}

/** settings シートを用意してアクティブにする（メニュー用） */
function openSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSettingsSheet(ss);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'C列の値を変えると次回更新から反映されます（株価上限の目安: '
    + getMaxStockPrice(getSettings(ss)).toLocaleString() + '円）', 'settings', 5);
}

function create_pickup_stocks(){
  runPickup(false);
}

/**
 * ネット取得を一切せず、キャッシュ済みの生データに現在の settings だけを
 * 適用して pickup を作り直す。settings を変えた直後の再判定用。
 * キャッシュは最大1時間で消えるため、それを過ぎるとデータなしとして案内する。
 */
function create_pickup_stocks_from_cache(){
  runPickup(true);
}

function runPickup(cacheOnly){

  // キャッシュを有効化（開発時はfalseに設定可能）
  const USE_CACHE = true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  Logger.log('適用設定: ' + JSON.stringify(settings)
    + ' / 株価上限の目安: ' + getMaxStockPrice(settings) + '円');
  
  // プログレスバー用のHTMLを表示
  showProgress('データ取得中...', 0);
  
  const iList = create_industrycodelist(USE_CACHE)

  if (iList == null || iList.length === 0){
    const errorMsg = "業種リスト取得エラー: データを取得できませんでした";
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, 'エラー', 10);
    return;
  }

  let stockTable = []
  const totalIndustries = iList.length;
  let processedIndustries = 0;
  let cacheMissPages = 0;

  for ( let industry of iList){

      for ( let pageNo = 1; pageNo < CONFIG.MAX_PAGES; pageNo++ ){

        const result = cacheOnly
          ? getCachedStocksByPage(industry.CODE, pageNo, settings)
          : get_stocks_by_page(industry.CODE, pageNo, USE_CACHE, settings)

        if (result.cacheMiss) cacheMissPages++;
        if ( result.done ) break  // そのページにデータなし = 最終ページ超え

        stockTable = stockTable.concat(result.stocks)
      }

      // 進捗を更新
      processedIndustries++;
      const progress = Math.round((processedIndustries / totalIndustries) * 100);
      showProgress('業種 ' + processedIndustries + '/' + totalIndustries + ' 処理中...', progress);
  }

  Logger.log('取得した銘柄数: ' + stockTable.length
    + (cacheOnly ? '（キャッシュのみ、ミス ' + cacheMissPages + 'ページ）' : ''));

  if (stockTable.length === 0) {
    const errorMsg = cacheOnly
      ? "キャッシュにデータがありません。「銘柄データ更新」を先に実行してください（キャッシュは最大1時間で消えます）"
      : "条件に合う銘柄が見つかりませんでした";
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, '警告', 10);
    return;
  }
  if (cacheOnly && cacheMissPages > 0) {
    Logger.log('注意: キャッシュ切れのページが ' + cacheMissPages + ' 件あり、結果が一部欠けている可能性があります');
  }

  // 第1.5段階: 25日平均出来高による流動性フィルタ（候補だけ個別取得する）
  const beforeVolume = stockTable.length;
  stockTable = applyVolumeFilter(stockTable, settings, USE_CACHE, (done, total) => {
    showProgress('出来高確認中 ' + done + '/' + total + '...', 90);
  }, cacheOnly);
  Logger.log('出来高フィルタ: ' + beforeVolume + '件 → ' + stockTable.length + '件');

  if (stockTable.length === 0) {
    const errorMsg = "出来高条件に合う銘柄が見つかりませんでした";
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, '警告', 10);
    return;
  }

  showProgress('スプレッドシートに書き込み中...', 95);

  // 割安な順（ミックス係数の小さい順）。同点はコード順（英字付きコードでも崩れないよう文字列比較）
  stockTable.sort((a, b) => (a.MIX - b.MIX) || String(a.TICKER).localeCompare(String(b.TICKER)));

  // 出力先は pickup シートに固定する。アクティブシートを消すと settings を
  // 開いたまま更新したときに設定ごと消えるため、getActiveSheet は使わない。
  let sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET_NAME);
  }

  // シートをクリア
  sheet.clear()
  Logger.log('シートをクリアしました')
  Logger.log('出力シート名: ' + sheet.getName());

  // ヘッダーとデータを配列にまとめる
  let header = ['コード','名称','株価','前日比','騰落率','PER','PBR','配当利回','ミックス係数','購入金額（単元）','出来高（25日平均）']
  let allData = [header]

  for( let stock of stockTable ){
    const rowData = [stock.TICKER, stock.NAME, stock.PRICE, stock.CHANGE, stock.CHANGE_RATE, stock.PER, stock.PBR, stock.DPS,
      stock.MIX, Number(stock.PRICE) * settings.unitShares,
      (stock.AVG_VOLUME === null || stock.AVG_VOLUME === undefined) ? '不明' : stock.AVG_VOLUME]
    allData.push(rowData)
  }

  Logger.log('書き込むデータ行数: ' + allData.length);
  Logger.log('最初の3行のデータ: ' + JSON.stringify(allData.slice(0, 3)));

  // 一括書き込み
  try {
    if (allData.length > 1) {
      const dataRange = sheet.getRange(1, 1, allData.length, header.length);
      dataRange.setValues(allData);
      
      // フォントサイズを13に設定
      dataRange.setFontSize(13);
      
      // コード列（A列）を右詰めに設定
      sheet.getRange(1, 1, allData.length, 1).setHorizontalAlignment('right');
      
      // 前日比（D列）と騰落率（E列）を右詰めに設定
      sheet.getRange(1, 4, allData.length, 2).setHorizontalAlignment('right');
      
      // 名称列（B列）にYahoo!ファイナンスへのリンクを一括設定する
      const nameFormulas = stockTable.map((stock) => [createYahooFinanceFormula(stock.TICKER, stock.NAME)]);
      sheet.getRange(2, 2, stockTable.length, 1).setFormulas(nameFormulas);
      
      Logger.log('スプレッドシートに書き込み完了');
    } else {
      Logger.log('書き込むデータがありません');
    }
  } catch(e) {
    const errorMsg = 'スプレッドシート書き込みエラー: ' + e.toString();
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, 'エラー', 10);
    throw e;
  }
  
  // 最終更新日時を記録（ヘッダーが K 列まで伸びたため M 列へ）
  const lastUpdateCell = sheet.getRange('M1');
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  lastUpdateCell.setValue('最終更新: ' + formattedDate);
  lastUpdateCell.setFontSize(9);
  lastUpdateCell.setFontColor('#666666');
  
  // 完了メッセージを表示
  closeProgress();

}

/**
 * Yahoo!ファイナンスへのリンク数式を作成
 * 銘柄名に含まれるダブルクォートは数式の文字列としてエスケープする
 */
function createYahooFinanceFormula(ticker, name) {
  const escapedName = String(name).replace(/"/g, '""');
  const url = 'https://finance.yahoo.co.jp/quote/' + encodeURIComponent(ticker) + '.T';
  return '=HYPERLINK("' + url + '","' + escapedName + '")';
}

/**
 * 進捗をトースト通知で表示
 */
function showProgress(message, percent) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message + ' (' + percent + '%)', '銘柄データ取得中', 3);
}

/**
 * 完了通知
 */
function closeProgress() {
  SpreadsheetApp.getActiveSpreadsheet().toast('データ取得が完了しました', '完了', 5);
}

/**
 * 株探内の業種名リストの作成
 * 
 * 株探のTOPIX33業種名を固定リストで返す。
 * 新URL: https://s.kabutan.jp/sectors/業種名/?page=N
 * 
 * （旧: https://kabutan.jp/warning/?mode=9_1 からスクレイピング → JS動的レンダリングで取得不可）
 * （旧: https://kabutan.jp/themes/?industry=X → 404で廃止）
 */
function create_industrycodelist(useCache = true){

  // TOPIX33業種の固定リスト（s.kabutan.jp/sectors/NAME/ のNAMEに対応）
  const INDUSTRY_LIST = [
    {CODE: '水産・農林業',      NAME: '水産・農林業'},
    {CODE: '鉱業',              NAME: '鉱業'},
    {CODE: '建設業',            NAME: '建設業'},
    {CODE: '食料品',            NAME: '食料品'},
    {CODE: '繊維製品',          NAME: '繊維製品'},
    {CODE: 'パルプ・紙',        NAME: 'パルプ・紙'},
    {CODE: '化学',              NAME: '化学'},
    {CODE: '医薬品',            NAME: '医薬品'},
    {CODE: '石油・石炭',    NAME: '石油・石炭製品'},
    {CODE: 'ゴム製品',          NAME: 'ゴム製品'},
    {CODE: 'ガラス・土石',  NAME: 'ガラス・土石製品'},
    {CODE: '鉄鋼',              NAME: '鉄鋼'},
    {CODE: '非鉄金属',          NAME: '非鉄金属'},
    {CODE: '金属製品',          NAME: '金属製品'},
    {CODE: '機械',              NAME: '機械'},
    {CODE: '電気機器',          NAME: '電気機器'},
    {CODE: '輸送用機器',        NAME: '輸送用機器'},
    {CODE: '精密機器',          NAME: '精密機器'},
    {CODE: 'その他製品',        NAME: 'その他製品'},
    {CODE: '電気・ガス',      NAME: '電気・ガス業'},
    {CODE: '陸運業',            NAME: '陸運業'},
    {CODE: '海運業',            NAME: '海運業'},
    {CODE: '空運業',            NAME: '空運業'},
    {CODE: '倉庫・運輸',  NAME: '倉庫・運輸関連業'},
    {CODE: '情報・通信業',      NAME: '情報・通信業'},
    {CODE: '卸売業',            NAME: '卸売業'},
    {CODE: '小売業',            NAME: '小売業'},
    {CODE: '銀行業',            NAME: '銀行業'},
    {CODE: '証券・商品',            NAME: '証券・商品先物取引業'},
    {CODE: '保険業',            NAME: '保険業'},
    {CODE: 'その他金融業',      NAME: 'その他金融業'},
    {CODE: '不動産業',          NAME: '不動産業'},
    {CODE: 'サービス業',        NAME: 'サービス業'},
  ];

  Logger.log('業種リストを固定値から取得: ' + INDUSTRY_LIST.length + '件');
  return INDUSTRY_LIST;
}

/**
 * 業種名でURLエンコードして銘柄一覧を取得
 * 
 * 新URL: https://s.kabutan.jp/sectors/業種名/?page=N
 * 1ページ20件、?page=N でページネーション
 * 
 * 戻り値: { stocks: [...], done: bool }
 *   stocks: フィルタ条件を満たす銘柄リスト
 *   done: true = このページにデータなし → ループ終了すべき
 *
 * キャッシュにはフィルタ前の生データを保存する。settings を変えたら
 * キャッシュクリアなしで次回実行から即反映される。
 */
function get_stocks_by_page(industryCode, pageNo, useCache = true, settings){

  const s = settings || DEFAULT_SETTINGS;
  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('stocks', encodeURIComponent(industryCode), pageNo);

  // キャッシュから生データを取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      const raw = JSON.parse(cached);
      return { stocks: applyScreening(raw.stocks, s), done: raw.done };
    }
  }

  const encodedIndustry = encodeURIComponent(industryCode);
  const url = 'https://s.kabutan.jp/sectors/' + encodedIndustry + '/?page=' + pageNo;
  let raw = { stocks: [], done: false };

  try {
    const text = fetchWithUserAgent(url);

    // <tbody>〜</tbody> を抽出
    const tbodyMatch = text.match(/<tbody>([\s\S]*?)<\/tbody>/);
    if (!tbodyMatch) {
      raw.done = true;
      cache.put(cacheKey, JSON.stringify(raw), CONFIG.CACHE_STOCK_DURATION);
      return { stocks: [], done: true };
    }

    const tbody = tbodyMatch[1];
    const trMatches = tbody.match(/<tr>([\s\S]*?)<\/tr>/g);
    if (!trMatches) {
      raw.done = true;
      cache.put(cacheKey, JSON.stringify(raw), CONFIG.CACHE_STOCK_DURATION);
      return { stocks: [], done: true };
    }

    let parsedCount = 0;
    for (const row of trMatches) {
      const stockData = get_stock_data(row);
      if (stockData && stockData.TICKER) {
        parsedCount++;
        raw.stocks.push(stockData);
      }
    }

    // ティッカーが1件も取れなかった = ページ範囲外
    if (parsedCount === 0) {
      raw.done = true;
    }

    cache.put(cacheKey, JSON.stringify(raw), CONFIG.CACHE_STOCK_DURATION);

  } catch(e) {
    Logger.log('銘柄ページ取得エラー (業種:' + industryCode + '/ページ' + pageNo + '): ' + e.toString());
    return { stocks: [], done: true };  // エラー時もループを止める（キャッシュしない）
  }

  return { stocks: applyScreening(raw.stocks, s), done: raw.done };
}

/**
 * キャッシュ済みの生データだけから1ページ分を返す（ネット取得はしない）。
 * キャッシュに無ければ cacheMiss: true を返し、その業種は打ち切る。
 * 戻り値の形は get_stocks_by_page と揃える。
 */
function getCachedStocksByPage(industryCode, pageNo, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('stocks', encodeURIComponent(industryCode), pageNo);
  const cached = cache.get(cacheKey);
  if (cached == null) {
    return { stocks: [], done: true, cacheMiss: true };
  }
  const raw = JSON.parse(cached);
  return { stocks: applyScreening(raw.stocks, s), done: raw.done, cacheMiss: false };
}

/**
 * HTMLの<tr>要素から個別銘柄情報を連想配列に抽出
 * 
 * s.kabutan.jp/sectors/業種名/ のHTML構造:
 *   <th> 銘柄名・コード (href="/stocks/TICKER/")
 *   <td> 株価
 *   <td> 前日比（値と率）
 *   <td> PER
 *   <td> PBR
 *   <td> 利回り
 */
function get_stock_data(row) {

  let res = {};

  // ティッカーを抽出: href="/stocks/TICKER/"
  const tickerMatch = row.match(/href="\/stocks\/([^"\/]+)\//);
  if (!tickerMatch) return res;
  res.TICKER = tickerMatch[1];

  // 銘柄名を抽出: <abbr title="正式名称"> または <p class="...">テキスト</p>
  const abbrMatch = row.match(/<abbr title="([^"]+)">/);
  if (abbrMatch) {
    res.NAME = abbrMatch[1];
  } else {
    const nameMatch = row.match(/<p class="font-bold[^"]*">([^<]+)<\/p>/);
    res.NAME = nameMatch ? nameMatch[1].trim() : res.TICKER;
  }

  // <td> を順番に抽出（株価、前日比、PER、PBR、利回り）
  const tdMatches = row.match(/<td>([\s\S]*?)<\/td>/g);
  if (!tdMatches || tdMatches.length < 5) return {};

  // 株価（カンマ除去）
  res.PRICE = removeHtmlTags(tdMatches[0]).replace(/,/g, '').trim();

  // 前日比: <span class='plus-num'>+129</span> or <span class='minus-num'>-10</span> or 0
  const changeMatch = row.match(/<span class='(?:plus|minus)-num'>([^<]+)<\/span>/);
  res.CHANGE = changeMatch ? changeMatch[1].trim() : '0';

  // 騰落率: 2番目の plus-num/minus-num span（<br> 後）
  const changeRateMatch = row.match(/<span class='text-2xs (?:plus-num|minus-num|)'>([^<]+)<\/span>/);
  if (changeRateMatch) {
    res.CHANGE_RATE = changeRateMatch[1].trim() + '%';
  } else {
    // 前日比ゼロの場合
    const zeroRateMatch = tdMatches[1].match(/(\d+\.\d+)/);
    res.CHANGE_RATE = zeroRateMatch ? zeroRateMatch[1] + '%' : '0%';
  }

  // PER（"倍"を除去、"ー"は0扱い）
  const perText = removeHtmlTags(tdMatches[2]).replace(/倍/g, '').trim();
  res.PER = (perText === 'ー' || perText === '') ? 0 : parseFloat(perText);

  // PBR
  const pbrText = removeHtmlTags(tdMatches[3]).replace(/倍/g, '').trim();
  res.PBR = (pbrText === 'ー' || pbrText === '') ? 0 : parseFloat(pbrText);

  // 配当利回り（"%"を除去）
  const dpsText = removeHtmlTags(tdMatches[4]).replace(/%/g, '').trim();
  res.DPS = (dpsText === 'ー' || dpsText === '') ? 0 : parseFloat(dpsText);

  // グレアムのミックス係数（PER・PBRのどちらかが欠損した場合は判定不能として null）
  res.MIX = (res.PER > 0 && res.PBR > 0) ? Math.round(res.PER * res.PBR * 1000) / 1000 : null;

  return res;
}