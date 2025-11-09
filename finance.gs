//deploytest のためのコメント

function testgetp() {
  let price;

  price = getprice_kabutan(8001);

  Logger.log(price);
}
function getper_kabutan(ticker_TSE) {
  // 株探から価格を取得する際のURLを、仮引数のticker_TSEを使って作る
  // 日本株で使用

  const url = 'https://kabutan.jp/stock/?code=' + ticker_TSE;
  const start_tag = '<span class="kabuka">';
  var neat_price = get_neatprice(url, start_tag);

  neat_price = neat_price.replace('円', '');

  return neat_price;
}

function getprice_kabutan(ticker_TSE) {
  // 株探から価格を取得する際のURLを、仮引数のticker_TSEを使って作る
  // 日本株で使用

  const url = 'http://kabutan.jp/stock/kabuka?code=' + ticker_TSE;
  const start_tag = '<span class="kabuka">';
  var neat_price = get_neatprice(url, start_tag);

  neat_price = neat_price.replace('円', '');

  return neat_price;
}

function getprice_yahoo(ticker) {
  // yahoo financeから価格を取得する際のURLを作成
  // 仮想通貨などgooglefinanceで対応していないものを取得するときに使用

  const url = 'https://finance.yahoo.com/quote/' + ticker;
  const start_tag = '<span class="Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)';

  var neat_price = get_neatprice(url, start_tag);

  const index = neat_price.indexOf('>');
  if (index != -1) {
    neat_price = neat_price.slice(index + 1, -1);
  }

  return neat_price;
}

function get_neatprice(url, start_tag) {
  // 上の関数たち(getprice_kabutan or getprice_yahoo)で作ったURLと価格部分の開始タグを受け取る
  // webサイトのソースをフェッチ、テキスト化、開始タグを見つけて価格部分を抽出する

  // web siteから抽出
  var fetched_html;
  try {
    fetched_html = UrlFetchApp.fetch(url);
  } catch (e) {
    return 'FetchError';
  }

  // htmlからテキストに変換
  const html_str = fetched_html.getContentText('UTF-8');
  if (html_str == '') {
    return 'PageEmpty';
  }

  // 開始タグから対応する終了タグを作成
  const splited_tag = start_tag.split(' ');
  var end_tag = splited_tag[0].slice(0, 1) + '/' + splited_tag[0].slice(1);
  if (end_tag.slice(-1) != '>') {
    end_tag += '>';
  }

  // webサイトのhtmlソースから、開始タグ〜終了タグで囲まれた部分(=価格)を抽出する
  const index_start = html_str.indexOf(start_tag);
  if (index_start == -1) return 'PriceEmpty';
  var price2tail = html_str.substring(index_start + start_tag.length); // price ... end tag ... end of html

  const index_end = price2tail.indexOf(end_tag);
  var price_neat = price2tail.substring(0, index_end); // price

  return price_neat;
}
