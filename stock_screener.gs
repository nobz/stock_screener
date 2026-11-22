/**
 * 設定定数
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

function create_pickup_stocks(){

  // キャッシュを有効化（開発時はfalseに設定可能）
  const USE_CACHE = true;
  
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

  for ( let industry of iList){

      for ( let pageNo = 1; pageNo < CONFIG.MAX_PAGES; pageNo++ ){

        const result = get_stocks_by_page(industry.CODE, pageNo, USE_CACHE)
       
        if ( result.done ) break  // そのページにデータなし = 最終ページ超え

        stockTable = stockTable.concat(result.stocks)
      }
      
      // 進捗を更新
      processedIndustries++;
      const progress = Math.round((processedIndustries / totalIndustries) * 100);
      showProgress('業種 ' + processedIndustries + '/' + totalIndustries + ' 処理中...', progress);
  }

  Logger.log('取得した銘柄数: ' + stockTable.length);
  
  if (stockTable.length === 0) {
    const errorMsg = "条件に合う銘柄が見つかりませんでした";
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, '警告', 10);
    return;
  }
  
  showProgress('スプレッドシートに書き込み中...', 95);

  stockTable.sort((a, b) => Number(a.TICKER) - Number(b.TICKER));

  const sheet = SpreadsheetApp.getActiveSheet()
  
  // シートをクリア
  sheet.clear()
  Logger.log('シートをクリアしました')
  Logger.log('アクティブシート名: ' + sheet.getName());

  // ヘッダーとデータを配列にまとめる
  let header = ['コード','名称','株価','前日比','騰落率','PER','PBR','配当利回']
  let allData = [header]

  for( let stock of stockTable ){
    const rowData = [stock.TICKER, stock.NAME, stock.PRICE, stock.CHANGE, stock.CHANGE_RATE, stock.PER, stock.PBR, stock.DPS]
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
  
  // 最終更新日時を記録（J列に移動）
  const lastUpdateCell = sheet.getRange('J1');
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
 */
function get_stocks_by_page(industryCode, pageNo, useCache = true){

  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('stocks', encodeURIComponent(industryCode), pageNo);
  
  // キャッシュから取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      return JSON.parse(cached);
    }
  }

  const encodedIndustry = encodeURIComponent(industryCode);
  const url = 'https://s.kabutan.jp/sectors/' + encodedIndustry + '/?page=' + pageNo;
  let result = { stocks: [], done: false };

  try {
    const text = fetchWithUserAgent(url);

    // <tbody>〜</tbody> を抽出
    const tbodyMatch = text.match(/<tbody>([\s\S]*?)<\/tbody>/);
    if (!tbodyMatch) {
      result.done = true;
      cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_STOCK_DURATION);
      return result;
    }

    const tbody = tbodyMatch[1];
    const trMatches = tbody.match(/<tr>([\s\S]*?)<\/tr>/g);
    if (!trMatches) {
      result.done = true;
      cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_STOCK_DURATION);
      return result;
    }

    let parsedCount = 0;
    for (const row of trMatches) {
      const stockData = get_stock_data(row);
      if (stockData && stockData.TICKER) {
        parsedCount++;
        // グレアムのミックス係数フィルタ（欠損時は判定不能として除外）
        if (stockData.MIX !== null && stockData.MIX < CONFIG.GRAHAM_MIX_THRESHOLD && stockData.DPS >= CONFIG.DIVIDEND_YIELD_THRESHOLD) {
          result.stocks.push(stockData);
        }
      }
    }

    // ティッカーが1件も取れなかった = ページ範囲外
    if (parsedCount === 0) {
      result.done = true;
    }

    cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_STOCK_DURATION);

  } catch(e) {
    Logger.log('銘柄ページ取得エラー (業種:' + industryCode + '/ページ' + pageNo + '): ' + e.toString());
    result.done = true;  // エラー時もループを止める
    return result;
  }

  return result;
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