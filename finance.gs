/**
 * 株探から株価を取得
 */
function getprice_kabutan(ticker_TSE) {
  try {
    const url = 'https://kabutan.jp/stock/kabuka?code=' + ticker_TSE;
    const start_tag = '<span class="kabuka">';
    let neat_price = get_neatprice(url, start_tag);

    if (neat_price === 'FetchError' || neat_price === 'PriceEmpty' || neat_price === 'PageEmpty') {
      Logger.log('株価取得失敗: ' + ticker_TSE + ' - ' + neat_price);
      return neat_price;
    }

    neat_price = neat_price.replace('円', '');

    return neat_price;
  } catch(e) {
    Logger.log('株価取得エラー: ' + ticker_TSE + ' - ' + e.toString());
    return 'Error';
  }
}

/**
 * HTMLから価格部分を抽出
 */
function get_neatprice(url, start_tag) {
  let fetched_html;
  try {
    fetched_html = UrlFetchApp.fetch(url);
  } catch (e) {
    Logger.log('価格取得エラー: ' + url + ' - ' + e.toString());
    return 'FetchError';
  }

  const html_str = fetched_html.getContentText('UTF-8');
  if (html_str === '') {
    Logger.log('空のページ: ' + url);
    return 'PageEmpty';
  }

  // 開始タグから対応する終了タグを作成
  const splited_tag = start_tag.split(' ');
  let end_tag = splited_tag[0].slice(0, 1) + '/' + splited_tag[0].slice(1);
  if (end_tag.slice(-1) !== '>') {
    end_tag += '>';
  }

  // 開始タグ〜終了タグで囲まれた部分を抽出
  const index_start = html_str.indexOf(start_tag);
  if (index_start === -1) {
    Logger.log('価格タグが見つかりません: ' + url);
    return 'PriceEmpty';
  }
  const price2tail = html_str.substring(index_start + start_tag.length);

  const index_end = price2tail.indexOf(end_tag);
  if (index_end === -1) {
    Logger.log('価格タグの終了タグが見つかりません: ' + url);
    return 'PriceEmpty';
  }
  const price_neat = price2tail.substring(0, index_end);
  if (price_neat === '') {
    Logger.log('価格が空です: ' + url);
    return 'PriceEmpty';
  }

  return price_neat;
}
