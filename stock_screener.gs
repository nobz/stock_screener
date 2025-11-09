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
  CACHE_CLEAR_MAX_PAGE: 20,
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
  const cache = CacheService.getScriptCache();
  
  // 全てのキャッシュを削除
  try {
    // Apps Scriptには全削除メソッドがないので、既知のキーを削除
    let keys = [getCacheKey('industry', 'list')];
    
    // 業種コード0-50、ページ1-20の範囲でクリア
    for (let i = 0; i <= CONFIG.CACHE_CLEAR_MAX_INDUSTRY; i++) {
      for (let p = 1; p <= CONFIG.CACHE_CLEAR_MAX_PAGE; p++) {
        keys.push(getCacheKey('stocks', i, p));
      }
    }
    
    // 100個ずつ削除（removeAllの制限対策）
    for (let i = 0; i < keys.length; i += CONFIG.CACHE_BATCH_SIZE) {
      const batch = keys.slice(i, i + CONFIG.CACHE_BATCH_SIZE);
      cache.removeAll(batch);
    }
    
    Logger.log('キャッシュを完全にクリアしました: ' + keys.length + '個のキー');
  } catch(e) {
    const errorMsg = 'キャッシュクリアエラー: ' + e.toString();
    Logger.log(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, 'エラー', 10);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('キャッシュをクリアしました', '完了', 3);
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

        const stocks = get_stocks_by_page(industry.CODE, pageNo, USE_CACHE)
       
        if ( stocks.length == 0) break

        stockTable = stockTable.concat(stocks)
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
  let header = ['コード','名称','PER','PBR','配当利回']
  let allData = [header]

  for( let stock of stockTable ){
    const rowData = [stock.TICKER, stock.NAME, stock.PER, stock.PBR, stock.DPS]
    allData.push(rowData)
  }

  Logger.log('書き込むデータ行数: ' + allData.length);
  Logger.log('最初の3行のデータ: ' + JSON.stringify(allData.slice(0, 3)));

  // 一括書き込み
  try {
    if (allData.length > 1) {
      sheet.getRange(1, 1, allData.length, header.length).setValues(allData)
      
      // コード列（A列）を右詰めに設定
      sheet.getRange(1, 1, allData.length, 1).setHorizontalAlignment('right');
      
      // 名称列（B列）にYahoo!ファイナンスへのリンクを設定
      for (let i = 0; i < stockTable.length; i++) {
        const ticker = stockTable[i].TICKER;
        const name = stockTable[i].NAME;
        const url = 'https://finance.yahoo.co.jp/quote/' + ticker + '.T';
        const nameFormula = '=HYPERLINK("' + url + '","' + name + '")';
        sheet.getRange(i + 2, 2).setFormula(nameFormula);
      }
      
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
  
  // 完了メッセージを表示
  closeProgress();

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
 * 株探内の業種IDリストの作成
 * 
 * 例えば海運の銘柄一覧は
 * 　https://kabutan.jp/themes/?industry=22&market=1
 * このindustryのコードを全33業種分取得する
 * 
 * 一覧ページから解析する
 *  https://kabutan.jp/warning/?mode=9_1
 * 
 * 
 */
function create_industrycodelist(useCache = true){

  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('industry', 'list');
  
  // キャッシュから取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      Logger.log('業種リストをキャッシュから取得');
      return JSON.parse(cached);
    }
  }

  Logger.log('業種リストを新規取得');
  let iList = []

  for (let pageNo=1; pageNo < CONFIG.MAX_PAGES; pageNo++){
    
    const industryList = get_industry_by_page(pageNo)

    if( _list.length == 0 ) break;

    iList = iList.concat(_list)
  }

  iList.sort((a, b) => Number(a.CODE) - Number(b.CODE));

  // キャッシュに保存
  cache.put(cacheKey, JSON.stringify(iList), CONFIG.CACHE_INDUSTRY_DURATION);

  return iList
}
function get_industry_by_page(pageNo){

  let iList = []

  const url = 'https://kabutan.jp/warning/?mode=9_1&page=' + pageNo

  try {
    const text = fetchWithUserAgent(url);
  
    const stock_table = Parser.data(text).from('<table class="stock_table">').to('/table>').build();
    const tr_datas = Parser.data(stock_table).from('<tr>').to('</tr>').iterate();
    
    for ( const row of tr_datas){
      const industryData = get_industry_data(row)
      if (industryData == null) continue
      iList.push(industryData)
    }
  } catch(e){
    Logger.log('業種ページ取得エラー: ' + e.toString());
  }

  return iList
}

function get_stocks_by_page(industryNo, pageNo, useCache = true){

  let res = []

  const cache = CacheService.getScriptCache();
  const cacheKey = getCacheKey('stocks', industryNo, pageNo);
  
  // キャッシュから取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      return JSON.parse(cached);
    }
  }

  const url = 'https://kabutan.jp/themes/?industry='+industryNo + '&page='+pageNo

  try {
    const text = fetchWithUserAgent(url);
 
    const stock_table = Parser.data(text).from('<table class="stock_table st_market">').to('/table>').build();

    const tr_datas = Parser.data(stock_table).from('<tr>').to('</tr>').iterate();

    for ( const row of tr_datas){

      const stockData = get_stock_data(row)

      if ( Object.keys(stockData).length ){

        // グレアムのミックス係数
        if( stockData.MIX < CONFIG.GRAHAM_MIX_THRESHOLD && stockData.DPS >= CONFIG.DIVIDEND_YIELD_THRESHOLD){
          res.push(stockData)
        }
      }
    }

    // キャッシュに保存
    cache.put(cacheKey, JSON.stringify(res), CONFIG.CACHE_STOCK_DURATION);

  } catch(e) {
    Logger.log('銘柄ページ取得エラー (業種' + industryNo + '/ページ' + pageNo + '): ' + e.toString());
    return res;
  }

  return res
}

/**
 * HTMLの文字列から個別銘柄情報を連想配列に抽出
 */
function get_stock_data( td ){

  let res = {}

  if (td.includes('<td class="tac">') == false) return res;

  //let datas = Parser.data(td).from('<td>').from('</td>').iterate();
 
  let datas = td.split('\r\n')

  res.TICKER = removeHtmlTags(datas[1])
  res.NAME = removeHtmlTags(datas[2])
  res.PER = removeHtmlTags(datas[11])
  res.PBR = removeHtmlTags(datas[12])
  res.DPS = removeHtmlTags(datas[13])

  res.MIX = Math.round(res.PER * res.PBR*1000)/1000

  return res
}


/**
 * HTMLの文字列から業種情報を連想配列に抽出
 * 
 * 
 * 
  <tr>
  <td class="tac"><a href="/stock/?code=0256">0256</a></td>
  <th scope="row" class="tal"><a href="/themes/?industry=6&market=1">パルプ・紙</a></th>
  <td>10</td>
  <td class="chart_icon"><a href="/stock/chart?code=0256"></a></td>
  <td>533.08</td>
  <td></td>
  <td class="w61"><span class="up">+7.60</span></td>
  <td class="w50"><span class="up">+1.45</span>%</td>
  <td>13.8</td>
  <td>0.59</td>
  <td>2.32</td>
  </tr>
 * 
 * 
 * 
 * 
 * 
 */
function get_industry_data( td ){

  let res = {}

  if (td.includes('<td class="tac">') == false) return null;

  let datas = td.split('\r\n')

  // <th scope="row" class="tal"><a href="/themes/?industry=6&market=1">パルプ・紙</a></th>
  const matchData = datas[2].match(/(?<=industry=)[0-9]+/)  // industryの値を抽出

  if(matchData == null) return null;

  res.CODE = matchData[0]
  res.NAME = removeHtmlTags(datas[2])

  return res
}