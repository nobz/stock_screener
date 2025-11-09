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
    let keys = ['industry_list'];
    
    // 業種コード0-50、ページ1-20の範囲でクリア
    for (let i = 0; i <= 50; i++) {
      for (let p = 1; p <= 20; p++) {
        keys.push('stocks_' + i + '_' + p);
      }
    }
    
    // 100個ずつ削除（removeAllの制限対策）
    for (let i = 0; i < keys.length; i += 100) {
      const batch = keys.slice(i, i + 100);
      cache.removeAll(batch);
    }
    
    Logger.log('キャッシュを完全にクリアしました: ' + keys.length + '個のキー');
  } catch(e) {
    Logger.log('キャッシュクリアエラー: ' + e.toString());
  }
}

function create_pickup_stocks(){

  // キャッシュを有効化（開発時はfalseに設定可能）
  const USE_CACHE = true;
  
  const iList = create_industrycodelist(USE_CACHE)

  if (iList == null){
    Logger.log("業種リスト取得エラー")
  }

  let stockTable = new Array()

  for ( let _d of iList){

      for ( let pageNo = 1; pageNo<99; pageNo++ ){

        res = get_stocksByPage(_d.CODE, pageNo, USE_CACHE)
       
        if ( res.length == 0) break

        stockTable = stockTable.concat(res)
      }
  }

  Logger.log('取得した銘柄数: ' + stockTable.length);

  stockTable = stockTable.sort (function(a,b){
    _a = Number(a.TICKER)
    _b = Number(b.TICKER)
    return (_a < _b)? -1:1;
  })

  const sheet = SpreadsheetApp.getActiveSheet()
  
  // シートをクリア
  sheet.clear()
  Logger.log('シートをクリアしました')
  Logger.log('アクティブシート名: ' + sheet.getName());

  // ヘッダーとデータを配列にまとめる
  let header = ['コード','名称','PER','PBR','配当利回']
  let allData = [header]

  for( let _s of stockTable ){
    _sd = [_s.TICKER, _s.NAME, _s.PER, _s.PBR, _s.DPS]
    allData.push(_sd)
  }

  Logger.log('書き込むデータ行数: ' + allData.length);
  Logger.log('最初の3行のデータ: ' + JSON.stringify(allData.slice(0, 3)));

  // 一括書き込み
  if (allData.length > 1) {
    sheet.getRange(1, 1, allData.length, header.length).setValues(allData)
    
    // コード列（A列）を右詰めに設定
    sheet.getRange(1, 1, allData.length, 1).setHorizontalAlignment('right');
    
    // 名称列（B列）にYahoo!ファイナンスへのリンクを設定
    for (let i = 0; i < stockTable.length; i++) {
      const ticker = stockTable[i].TICKER;
      const name = stockTable[i].NAME;
      const url = 'https://finance.yahoo.co.jp/quote/' + ticker + '.T';
      const formula = '=HYPERLINK("' + url + '","' + name + '")';
      sheet.getRange(i + 2, 2).setFormula(formula);
    }
    
    Logger.log('スプレッドシートに書き込み完了');
  } else {
    Logger.log('書き込むデータがありません');
  }

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
  const cacheKey = 'industry_list';
  
  // キャッシュから取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      Logger.log('業種リストをキャッシュから取得');
      return JSON.parse(cached);
    }
  }

  Logger.log('業種リストを新規取得');
  let iList = new Array()

  for (let pageNo=1; pageNo<99; pageNo++){
    
    let _list = get_industryByPage(pageNo)

    if( _list.length == 0 ) break;

    iList = iList.concat(_list)
  }

  iList = iList.sort(function(a,b){
    _a = Number( a.CODE)
    _b = Number( b.CODE)
    return (_a < _b)? -1:1;
  })

  // キャッシュに保存（6時間）
  cache.put(cacheKey, JSON.stringify(iList), 21600);

  return iList
}
function get_industryByPage(pageNo){

  let iList = new Array()
  let fetched_html

  const url = 'https://kabutan.jp/warning/?mode=9_1&page=' + pageNo

  // User-Agentヘッダーを追加してリクエスト
  const options = {
    'headers': {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    },
    'muteHttpExceptions': true
  };

  try {
    fetched_html = UrlFetchApp.fetch(url, options);
    
    // ステータスコードをチェック
    const responseCode = fetched_html.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('Error: HTTP ' + responseCode + ' for URL: ' + url);
      return iList;
    }
  } catch(e){
    Logger.log('Fetch error: ' + e.toString());
    return iList
  }
  let text = fetched_html.getContentText("utf-8");
  
  const stock_table = Parser.data(text).from('<table class="stock_table">').to('/table>').build();

  const tr_datas = Parser.data(stock_table).from('<tr>').to('</tr>').iterate();
  
  for ( td of tr_datas){

    let d = getIndustryData(td)

    if (d == null) continue

    iList.push(d)
  }

  return iList
}

function get_stocksByPage(industryNo, pageNo, useCache = true){

  let res = new Array()

  const cache = CacheService.getScriptCache();
  const cacheKey = 'stocks_' + industryNo + '_' + pageNo;
  
  // キャッシュから取得を試みる
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached != null) {
      return JSON.parse(cached);
    }
  }

  const url = 'https://kabutan.jp/themes/?industry='+industryNo + '&page='+pageNo

  // User-Agentヘッダーを追加してリクエスト
  const options = {
    'headers': {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    
    // ステータスコードをチェック
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('Error: HTTP ' + responseCode + ' for URL: ' + url);
      return res;
    }
    
    let text = response.getContentText("utf-8");
 
    const stock_table = Parser.data(text).from('<table class="stock_table st_market">').to('/table>').build();

    const tr_datas = Parser.data(stock_table).from('<tr>').to('</tr>').iterate();

    for ( td of tr_datas){

      let d = getStokData(td)

      if ( Object.keys(d).length ){

        // グレアムのミックス係数
        if( d.MIX < 11 && d.DPS >= 3.5){
          res.push(d)
        }
      }
    }

    // キャッシュに保存（1時間）
    cache.put(cacheKey, JSON.stringify(res), 3600);

  } catch(e) {
    Logger.log('Error processing stocks: ' + e.toString());
  }

  return res
}

/**
 * HTMLの文字列から個別銘柄情報を連想配列に抽出
 */
function getStokData( td ){

  let res = {}

  if (td.includes('<td class="tac">') == false) return res;

  //let datas = Parser.data(td).from('<td>').from('</td>').iterate();
 
  let datas = td.split('\r\n')

  let _d
  _d = datas[1].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')
  res.TICKER = _d.replace('\r\n','')

  _d = datas[2].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')

  res.NAME = _d.replace('\r\n','')

  _d = datas[11].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')

  res.PER = _d.replace('\r\n','')

  _d = datas[12].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')

  res.PBR = _d.replace('\r\n','')

  _d = datas[13].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')

  res.DPS = _d.replace('\r\n','')

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
function getIndustryData( td ){

  let res = {}

  if (td.includes('<td class="tac">') == false) return null;

  let datas = td.split('\r\n')

  // <th scope="row" class="tal"><a href="/themes/?industry=6&market=1">パルプ・紙</a></th>
  let matchData = datas[2].match(/(?<=industry=)[0-9]+/)  // industryの値を抽出

  if(matchData == null) return null;

  res.CODE = matchData[0]

  let name = datas[2].replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'')

  res.NAME = name

  return res
}

function get_neatprice(url, start_tag){  
  // 上の関数たち(getprice_kabutan or getprice_yahoo)で作ったURLと価格部分の開始タグを受け取る
  // webサイトのソースをフェッチ、テキスト化、開始タグを見つけて価格部分を抽出する

  // web siteから抽出
  var fetched_html ;
  try {
    fetched_html = UrlFetchApp.fetch(url);
  } catch(e){
    return 'FetchError';
  }

  // htmlからテキストに変換
  const html_str = fetched_html.getContentText('UTF-8');
  if (html_str == '') {
    return 'PageEmpty';
  }

  // 開始タグから対応する終了タグを作成
  const splited_tag = start_tag.split(' ');
  var end_tag = splited_tag[0].slice(0,1) + '/' + splited_tag[0].slice(1);
  if (end_tag.slice(-1) != '>') {
    end_tag += '>';
  }

  // webサイトのhtmlソースから、開始タグ〜終了タグで囲まれた部分(=価格)を抽出する
  const index_start = html_str.indexOf(start_tag);
  if (index_start == -1) return 'PriceEmpty';
  var price2tail = html_str.substring(index_start + start_tag.length); // price ... end tag ... end of html 
   
  const index_end = price2tail.indexOf(end_tag);
  var price_neat = price2tail.substring(0,index_end);  // price

  return price_neat;
  
}