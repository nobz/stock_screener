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
const HISTORY_SHEET_NAME = 'history';

/** PBR・自己資本比率の水準が事業会社と異なるため、注意列で区別する業種（create_industrycodelist の CODE） */
const FINANCIAL_INDUSTRY_CODES = ['銀行業', '証券・商品', '保険業', 'その他金融業'];
const FINANCIAL_NOTE = '金融業（PBR・自己資本比率の水準が事業会社と異なる）';

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
  { key: 'maxPayoutRatio', label: '予想配当性向の上限', value: 70, description: '％。予想1株配÷予想1株益がこれを超えると pickup の注意列に書く（除外はしない）' },
  { key: 'maxNetToOrdinaryRatio', label: '最終益/経常益の上限（特別利益の目安）', value: 90, description: '％。予想最終益が予想経常益のこの割合を超える分は特別利益とみなし、最終益をここまで抑えた補正PERでミックス係数を計算する' },
];

const DEFAULT_SETTINGS = SETTINGS_DEFINITIONS.reduce((acc, def) => {
  acc[def.key] = def.value;
  return acc;
}, {});

/**
 * settings シートが無ければ作り、不足している項目行だけを末尾に足す。
 * 定義済みキーの設定名（B列）と説明（D列）は定義に合わせて更新するが、
 * 値（C列）は絶対に上書きしない。
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
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    rows.forEach((row, index) => {
      if (!row[0]) return;
      const key = String(row[0]);
      knownKeys.add(key);
      const def = SETTINGS_DEFINITIONS.find((d) => d.key === key);
      if (def && (row[1] !== def.label || row[3] !== def.description)) {
        const sheetRow = index + 2;
        sheet.getRange(sheetRow, 2).setValue(def.label);
        sheet.getRange(sheetRow, 4).setValue(def.description);
      }
    });
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
    // PC版は <time datetime=...>、スマホ版は <th class="sticky">11/20</th> が日付の行
    if (!/<time\s+datetime=/.test(row) && !/<th class="sticky">\s*\d{1,2}\/\d{1,2}\s*<\/th>/.test(row)) continue;
    const tdMatches = row.match(/<td>([\s\S]*?)<\/td>/g);
    if (!tdMatches || tdMatches.length < 7) continue;
    const volumeText = removeHtmlTags(tdMatches[tdMatches.length - 1]).replace(/[,株]/g, '').trim();
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
  return getAverageVolumes([ticker], useCache, cacheOnly).get(ticker);
}

const UNKNOWN_VOLUME = { average: null, daysUsed: 0 };

/** 複数銘柄の平均出来高を Map(ticker → { average, daysUsed }) で返す。キャッシュのみモードではネット取得しない */
function getAverageVolumes(tickers, useCache = true, cacheOnly = false, onBatch) {
  const values = collectByTicker(tickers, {
    type: 'volume',
    useCache: useCache,
    cacheOnly: cacheOnly,
    duration: CONFIG.CACHE_STOCK_DURATION,
    // PC版（kabutan.jp/stock/...）は GAS からだと HTTP 403 になるため、業種一覧と同じスマホ版を使う
    urlOf: (ticker) => 'https://s.kabutan.jp/stocks/' + encodeURIComponent(ticker) + '/historical_prices/daily/',
    parse: (html) => {
      const volumes = parseDailyVolumes(html);
      return { average: computeAverageVolume(volumes, VOLUME_AVERAGE_DAYS), daysUsed: Math.min(volumes.length, VOLUME_AVERAGE_DAYS) };
    },
    shouldCache: (result) => result.average !== null,
    onBatch: onBatch,
  });
  const results = new Map();
  tickers.forEach((ticker) => results.set(ticker, values.get(ticker) || UNKNOWN_VOLUME));
  return results;
}

/** 同時に取得する件数。株探への負荷と速度の兼ね合い */
const FETCH_BATCH_SIZE = 10;

/**
 * 銘柄ごとのページを、キャッシュはまとめて読み、足りない分だけ FETCH_BATCH_SIZE 件ずつ同時に取得する。
 * 1件ずつ順番に取ると約1,000件で15分以上かかるため。
 * 戻り値: Map(ticker → parse の結果)。取得・解析できなかった銘柄は入らない。
 * onBatch(done, total) はバッチごとに呼ぶ（進捗表示と時間予算の確認用）。
 */
function collectByTicker(tickers, options) {
  const cache = CacheService.getScriptCache();
  const values = new Map();
  const failures = options.failures || new Map();
  const total = tickers.length;
  for (let start = 0; start < total; start += FETCH_BATCH_SIZE) {
    const chunk = tickers.slice(start, start + FETCH_BATCH_SIZE);
    const keys = chunk.map((ticker) => getCacheKey(options.type, ticker));

    if (options.useCache || options.cacheOnly) {
      const cached = cache.getAll(keys);
      chunk.forEach((ticker, i) => {
        if (cached[keys[i]] != null) values.set(ticker, JSON.parse(cached[keys[i]]));
      });
    }

    const missing = options.cacheOnly ? [] : chunk.filter((ticker) => !values.has(ticker));
    if (missing.length > 0) {
      const fetched = fetchAllWithUserAgent(missing.map(options.urlOf));
      const toCache = {};
      missing.forEach((ticker, i) => {
        if (fetched[i].text == null) {
          failures.set(ticker, fetched[i].reason);
          return;
        }
        try {
          const value = options.parse(fetched[i].text);
          if (value == null) {
            failures.set(ticker, '表なし');
            return;
          }
          values.set(ticker, value);
          if (!options.shouldCache || options.shouldCache(value)) {
            toCache[getCacheKey(options.type, ticker)] = JSON.stringify(value);
          }
        } catch (e) {
          failures.set(ticker, '解析エラー');
          Logger.log(options.type + ' 解析エラー (' + ticker + '): ' + e.toString());
        }
      });
      if (Object.keys(toCache).length > 0) cache.putAll(toCache, options.duration);
    }

    if (options.onBatch) options.onBatch(Math.min(start + FETCH_BATCH_SIZE, total), total);
  }
  return values;
}

/**
 * 複数URLを同時に取得し、[{ text, reason }] を返す。失敗時は text: null で reason に理由（"HTTP 403" など）。
 * fetchAll は1件でも通信エラーがあると全体が例外になるため、そのときだけ1件ずつ取り直す。
 */
function fetchAllWithUserAgent(urls) {
  const requests = urls.map((url) => ({
    url: url,
    headers: { 'User-Agent': CONFIG.USER_AGENT },
    muteHttpExceptions: true,
  }));
  try {
    return UrlFetchApp.fetchAll(requests).map((response, i) => {
      const code = response.getResponseCode();
      if (code === 200) return { text: response.getContentText('utf-8'), reason: null };
      Logger.log('HTTP ' + code + ': ' + urls[i]);
      return { text: null, reason: 'HTTP ' + code };
    });
  } catch (e) {
    Logger.log('まとめて取得に失敗したため1件ずつ取り直します: ' + e.toString());
    return urls.map((url) => {
      try {
        return { text: fetchWithUserAgent(url), reason: null };
      } catch (err) {
        const match = String(err).match(/HTTP \d+/);
        return { text: null, reason: match ? match[0] : '通信エラー' };
      }
    });
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
  const list = stocks || [];
  const volumes = getAverageVolumes(list.map((stock) => stock.TICKER), useCache, cacheOnly, onProgress);
  list.forEach((stock) => {
    const result = volumes.get(stock.TICKER);
    stock.AVG_VOLUME = result.average;
    if (result.average !== null && result.average <= s.minAverageVolume) {
      dropped.push(stock.TICKER);
    } else {
      kept.push(stock);
    }
  });
  if (dropped.length > 0) {
    Logger.log('出来高フィルタで除外: ' + dropped.length + '件 (' + dropped.join(', ') + ')');
  }
  return kept;
}

/** 財務の判定結果は業績が日次で変わらないため長めに持つ（CacheService の上限は6時間） */
const CACHE_FINANCE_DURATION = 21600;

/**
 * 財務キャッシュの種別。キャッシュには読み取り済みの業績行を入れるため、
 * 読む項目を変えたら番号を上げる（v2: 経常益を追加）
 */
const FINANCE_CACHE_TYPE = 'finance_v2';

/** "1,234" "52*" "－" などを数値化する。読めなければ null */
function parseFinanceNumber(text) {
  const cleaned = removeHtmlTags(String(text || '')).replace(/&nbsp;/g, '').replace(/[,*\s円]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  return Number(cleaned);
}

/**
 * 財務ページの通期業績表（ヘッダーに「修正1株配」がある最初の表）を読む。
 * PC版の行: <th scope="row"><span class="kubun1">I 予 </span>2027.03</th>
 * スマホ版の行: <th class="sticky">予 2027.03 <span ...></span></th>
 *     どちらも <td>売上高</td><td>営業益</td><td>経常益</td><td>最終益</td><td>1株益</td><td>1株配</td>... の順
 * 戻り値: [{ period, isForecast, ordinaryIncome, netIncome, eps, dps }]（ページ上の順＝古い順）
 * IFRS の会社は経常益の列に税引前利益が入る。
 */
function parseAnnualResults(html) {
  const text = String(html || '');
  const pcHeader = text.indexOf('data-help="修正1株配"');
  const headerIndex = pcHeader !== -1 ? pcHeader : text.indexOf('修正<br>1株配');
  if (headerIndex === -1) return [];
  const tableEnd = text.indexOf('</table>', headerIndex);
  const table = text.substring(headerIndex, tableEnd === -1 ? text.length : tableEnd);

  const rows = [];
  const trMatches = table.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) || [];
  for (const tr of trMatches) {
    const thMatch = tr.match(/<th (?:scope="row"|class="sticky")[^>]*>([\s\S]*?)<\/th>/);
    if (!thMatch) continue;
    const periodMatch = thMatch[1].match(/(\d{4}\.\d{2})/);
    if (!periodMatch) continue;
    const tds = tr.match(/<td[^>]*>[\s\S]*?<\/td>/g) || [];
    if (tds.length < 6) continue;
    rows.push({
      period: periodMatch[1],
      isForecast: thMatch[1].indexOf('予') !== -1,
      ordinaryIncome: parseFinanceNumber(tds[2]),
      netIncome: parseFinanceNumber(tds[3]),
      eps: parseFinanceNumber(tds[4]),
      dps: parseFinanceNumber(tds[5]),
    });
  }
  return rows;
}

/**
 * 配当の持続性を判定する。除外はせず、注意として出す理由の配列を返す。
 * 下方修正の履歴は株探プレミアム限定のため、「今期予想が前期実績より減益か」で代用する。
 */
function evaluateDividendSafety(rows, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const flags = [];
  const actuals = (rows || []).filter((r) => !r.isForecast);
  const forecasts = (rows || []).filter((r) => r.isForecast);
  const forecast = forecasts.length > 0 ? forecasts[forecasts.length - 1] : null;
  const lastActual = actuals.length > 0 ? actuals[actuals.length - 1] : null;
  let payoutRatio = null;

  if (forecast && forecast.eps !== null) {
    if (forecast.eps <= 0) {
      flags.push('予想赤字');
    } else if (forecast.dps !== null) {
      payoutRatio = Math.round(forecast.dps / forecast.eps * 1000) / 10;
      if (payoutRatio > s.maxPayoutRatio) {
        flags.push('配当性向 ' + payoutRatio + '%');
      }
    }
  }

  if (forecast && lastActual && forecast.dps !== null && lastActual.dps !== null
      && forecast.dps < lastActual.dps) {
    flags.push('減配予想');
  }

  for (let i = 1; i < actuals.length; i++) {
    if (actuals[i].dps !== null && actuals[i - 1].dps !== null && actuals[i].dps < actuals[i - 1].dps) {
      flags.push(actuals[i].period + ' 減配');
    }
  }

  if (forecast && lastActual && forecast.netIncome !== null && lastActual.netIncome !== null
      && lastActual.netIncome > 0 && forecast.netIncome < lastActual.netIncome) {
    const change = Math.round((forecast.netIncome / lastActual.netIncome - 1) * 100);
    flags.push('減益予想 ' + change + '%');
  }

  return { payoutRatio: payoutRatio, flags: flags };
}

/**
 * 特別利益による PER の見かけの低さを補正する。戻り値: { adjustedPer, ratio, flags[] }
 * 予想最終益が「予想経常益 × maxNetToOrdinaryRatio」を超える分を特別利益とみなし、
 * 最終益をそこまで抑えて PER を計算し直す: 補正PER = PER × max(1, 比率 ÷ 上限)。
 * 上限以下の会社は変えず、特別損失で最終益が小さい会社も持ち上げない
 * （そうした会社は第1段階の株探PERで既に落ちており、救うと基準が揃わないため）。
 * adjustedPer が null なら判定不能（経常赤字予想）。
 */
function evaluateEarningsQuality(rows, per, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const forecasts = (rows || []).filter((r) => r.isForecast);
  const forecast = forecasts.length > 0 ? forecasts[0] : null;
  const ordinary = forecast ? forecast.ordinaryIncome : null;
  const net = forecast ? forecast.netIncome : null;

  if (ordinary == null || net == null) {
    return { adjustedPer: per, ratio: null, flags: ['経常ベース判定不可'] };
  }
  if (net <= 0) {
    // 予想赤字は株探PERが出ず第1段階で落ちる。ここに来たら補正しない
    return { adjustedPer: per, ratio: null, flags: [] };
  }
  if (ordinary <= 0) {
    return { adjustedPer: null, ratio: null, flags: ['経常赤字予想'] };
  }

  const ratio = net / ordinary;
  const cap = s.maxNetToOrdinaryRatio / 100;
  if (ratio <= cap) {
    return { adjustedPer: per, ratio: ratio, flags: [] };
  }
  const adjustedPer = Math.round(per * ratio / cap * 100) / 100;
  return {
    adjustedPer: adjustedPer,
    ratio: ratio,
    flags: ['特別利益でかさ上げ（最終益/経常益 ' + Math.round(ratio * 100) + '%）'],
  };
}

/**
 * 候補に補正PERとミックス係数（補正後）を付け、閾値を満たさなくなった銘柄を除く。
 * 財務が無い銘柄は株探PERのまま残す（取得失敗で落とさない方針）。
 */
function applyEarningsQuality(stocks, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const kept = [];
  const dropped = [];
  (stocks || []).forEach((stock) => {
    const per = Number(stock.PER);
    let adjustedPer = per;
    if (stock.FINANCE_ROWS) {
      const result = evaluateEarningsQuality(stock.FINANCE_ROWS, per, s);
      adjustedPer = result.adjustedPer;
      stock.NOTES = (stock.NOTES || []).concat(result.flags);
    }
    stock.ADJUSTED_PER = adjustedPer;
    stock.MIX = (adjustedPer !== null && stock.PBR > 0) ? Math.round(adjustedPer * stock.PBR * 1000) / 1000 : null;
    if (stock.MIX !== null && stock.MIX < s.grahamMixThreshold) {
      kept.push(stock);
    } else {
      dropped.push(stock.TICKER);
    }
  });
  if (dropped.length > 0) {
    Logger.log('補正ミックス係数で除外: ' + dropped.length + '件 (' + dropped.join(', ') + ')');
  }
  return kept;
}

/**
 * 銘柄の配当の持続性を取得する。戻り値: { payoutRatio, flags[] }
 * パース後の判定材料（業績行）だけを銘柄単位で6時間キャッシュする。
 * 判定は settings に依存するため、キャッシュからでも毎回やり直す。
 * clearAllCache の対象外（キーを列挙できないため）。
 */
function getDividendSafety(ticker, settings, useCache = true, cacheOnly = false) {
  return getDividendSafeties([ticker], settings, useCache, cacheOnly).get(ticker);
}

/** 複数銘柄の配当判定を Map(ticker → { payoutRatio, flags[] }) で返す */
function getDividendSafeties(tickers, settings, useCache = true, cacheOnly = false, onBatch) {
  const failures = new Map();
  const rowsByTicker = collectByTicker(tickers, {
    failures: failures,
    type: FINANCE_CACHE_TYPE,
    useCache: useCache,
    cacheOnly: cacheOnly,
    duration: CACHE_FINANCE_DURATION,
    urlOf: (ticker) => 'https://s.kabutan.jp/stocks/' + encodeURIComponent(ticker) + '/finance/',
    parse: (html) => {
      const rows = parseAnnualResults(html);
      return rows.length > 0 ? rows : null;
    },
    onBatch: onBatch,
  });
  const results = new Map();
  tickers.forEach((ticker) => {
    const rows = rowsByTicker.get(ticker);
    if (rows) {
      results.set(ticker, Object.assign(evaluateDividendSafety(rows, settings), { rows: rows }));
      return;
    }
    const flag = cacheOnly ? '財務データ未取得' : '財務取得失敗（' + (failures.get(ticker) || '不明') + '）';
    results.set(ticker, { payoutRatio: null, flags: [flag], rows: null });
  });
  return results;
}

/** 候補全件に予想配当性向と注意を付ける（除外はしない） */
function applyDividendSafety(stocks, settings, useCache = true, onProgress, cacheOnly = false) {
  const list = stocks || [];
  const safeties = getDividendSafeties(list.map((stock) => stock.TICKER), settings, useCache, cacheOnly, onProgress);
  list.forEach((stock) => {
    const result = safeties.get(stock.TICKER);
    stock.PAYOUT_RATIO = result.payoutRatio;
    stock.FINANCE_ROWS = result.rows;
    stock.NOTES = (stock.IS_FINANCIAL ? [FINANCIAL_NOTE] : []).concat(result.flags);
  });
  return stocks;
}

/**
 * 業種一覧ページのキャッシュ種別。キャッシュには読み取り済みの値を入れるため、
 * 読み取り方を変えたら番号を上げて古いキャッシュを使わないようにする（v2: PER のカンマ修正）
 */
const STOCKS_CACHE_TYPE = 'stocks_v2';

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
    .addItem('銘柄データ更新', 'startPickup')
    .addItem('キャッシュで再判定（取得しない）', 'startPickupFromCache')
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
        keys.push(getCacheKey(STOCKS_CACHE_TYPE, encodeURIComponent(industry.CODE), p));
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

/**
 * メニュー用。サイドバーを開き、サイドバー側から runPickup を起動させる。
 * メニューから直接 runPickup を呼ぶと、終わるまでサイドバーを描画できないため。
 */
function startPickup() {
  openProgressSidebar(false);
}

function startPickupFromCache() {
  openProgressSidebar(true);
}

function openProgressSidebar(cacheOnly) {
  const template = HtmlService.createTemplateFromFile('progress');
  template.cacheOnly = cacheOnly;
  template.title = cacheOnly ? 'キャッシュで再判定' : '銘柄データ更新';
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('銘柄スクリーニング'));
}

/**
 * スクリプトエディタやトリガーから実行する用（進捗はサイドバーが無ければ誰も読まない）。
 * 時間予算で打ち切られた場合は { done: false } が返るので、もう一度実行すれば続きから進む。
 */
function create_pickup_stocks(){
  return runPickup(false);
}

/**
 * ネット取得を一切せず、キャッシュ済みの生データに現在の settings だけを
 * 適用して pickup を作り直す。settings を変えた直後の再判定用。
 * キャッシュは最大1時間で消えるため、それを過ぎるとデータなしとして案内する。
 */
function create_pickup_stocks_from_cache(){
  return runPickup(true);
}

/**
 * 候補を作って pickup に書く。戻り値は結果の要約（サイドバーの完了表示に使う）。
 * 候補0件などの終了理由も例外ではなく要約として返す。書き込み失敗は例外で返す。
 *
 * GAS の1回の実行は6分まで。時間予算を超えたら途中で打ち切り { done: false } を返す。
 * 取得済みのページ・出来高・財務はキャッシュに残っているため、呼び出し側（サイドバー）が
 * もう一度呼べば続きから進む。打ち切った回は pickup と history を書き換えない。
 * 戻り値: { done: boolean, message: string }
 */
function runPickup(cacheOnly, budgetMs, spreadsheet){
  Logger.log('runPickup 呼び出し (cacheOnly=' + !!cacheOnly + ')');
  // サイドバーとの通信が切れてもサーバー側の実行は続くことがある。
  // そこへ再試行が重なって pickup / history を二重に書かないよう、同時に1つだけ動かす。
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log('前の実行が動いているため待機を返します');
    return { done: false, busy: true, message: '前の実行がまだ動いています。終わるのを待っています' };
  }
  const startedAt = Date.now();
  const budget = (typeof budgetMs === 'number') ? budgetMs : TIME_BUDGET_MS;
  const checkBudget = () => {
    if (Date.now() - startedAt > budget) throw new TimeBudgetExceeded();
  };
  try {
    const message = runPickupCore(cacheOnly, checkBudget, startedAt, spreadsheet);
    deleteContinuationTriggers();
    return { done: true, message: message };
  } catch (e) {
    if (!(e instanceof TimeBudgetExceeded)) {
      deleteContinuationTriggers();
      throw e;
    }
    const message = '一旦区切りました。取得済みの分はキャッシュに残っており、続きから再開します';
    Logger.log(message + '（経過 ' + Math.round((Date.now() - startedAt) / 1000) + '秒）');
    writeProgress({ message: message, percent: 0, done: false });
    // サイドバーが続きを呼ばなくても進むよう、サーバー側でも続きを予約する。
    // サイドバーと競合した場合はロックで片方だけが動き、完了時に予約は消える。
    scheduleContinuation(cacheOnly);
    return { done: false, message: message };
  } finally {
    lock.releaseLock();
  }
}

/** 続きを実行するトリガーのハンドラ名と、発火までの待ち時間（1分） */
const CONTINUATION_HANDLER = 'continuePickupFromTrigger';
const CONTINUATION_DELAY_MS = 60000;
const CONTINUATION_SPREADSHEET_KEY = 'pickup_spreadsheet_id';
const CONTINUATION_CACHE_ONLY_KEY = 'pickup_cache_only';

/**
 * トリガー実行時はアクティブなスプレッドシートが無いため、通常実行のときに
 * 対象スプレッドシートの ID と実行モードをプロパティへ残しておく。
 */
function rememberContinuationTarget(ss, cacheOnly) {
  try {
    if (typeof PropertiesService === 'undefined' || !ss || !ss.getId) return;
    const props = PropertiesService.getScriptProperties();
    props.setProperty(CONTINUATION_SPREADSHEET_KEY, ss.getId());
    props.setProperty(CONTINUATION_CACHE_ONLY_KEY, cacheOnly ? '1' : '0');
  } catch (e) {
    Logger.log('継続先の保存エラー: ' + e.toString());
  }
}

/** 継続トリガーを1つだけ予約する（既存の継続トリガーは作り直す）。失敗してもサイドバー側の継続は残る */
function scheduleContinuation(cacheOnly) {
  try {
    if (typeof ScriptApp === 'undefined') return;
    deleteContinuationTriggers();
    ScriptApp.newTrigger(CONTINUATION_HANDLER).timeBased().after(CONTINUATION_DELAY_MS).create();
    Logger.log('続きを ' + Math.round(CONTINUATION_DELAY_MS / 1000) + '秒後にサーバー側で実行するよう予約しました');
  } catch (e) {
    Logger.log('継続トリガーの予約エラー: ' + e.toString());
  }
}

function deleteContinuationTriggers() {
  try {
    if (typeof ScriptApp === 'undefined') return;
    ScriptApp.getProjectTriggers()
      .filter((trigger) => trigger.getHandlerFunction() === CONTINUATION_HANDLER)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  } catch (e) {
    Logger.log('継続トリガーの削除エラー: ' + e.toString());
  }
}

/**
 * 時間トリガーから続きを実行する。対象は rememberContinuationTarget が残した
 * スプレッドシート ID で開く。実行中（busy）なら、動いている側が区切り時に
 * 予約し直す前提で、念のため自分でも予約し直す。
 */
function continuePickupFromTrigger() {
  let ss = null;
  let cacheOnly = false;
  try {
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty(CONTINUATION_SPREADSHEET_KEY);
    cacheOnly = props.getProperty(CONTINUATION_CACHE_ONLY_KEY) === '1';
    if (spreadsheetId) ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log('継続先スプレッドシートを開けませんでした: ' + e.toString());
  }
  if (!ss) {
    Logger.log('継続先スプレッドシートが不明なため、サーバー側の続きは実行しません');
    return;
  }
  const result = runPickup(cacheOnly, undefined, ss);
  if (result && result.busy) scheduleContinuation(cacheOnly);
}

/**
 * 1回の実行で使う時間の上限。GAS の6分制限より、サイドバーからの呼び出しが
 * 長いと通信が切れる方が先に起きるため、短く刻む。
 * 150秒でも区切りの応答がサイドバーへ届かず続きが呼ばれない事例があった
 * （2026-11-24）ため、90秒に縮めた。20回まで続けても最大で約30分であり、
 * 業種ページのキャッシュ（1時間）の範囲内に収まる。
 */
const TIME_BUDGET_MS = 90000;

function TimeBudgetExceeded() {
  this.name = 'TimeBudgetExceeded';
}

function runPickupCore(cacheOnly, checkBudget, startedAt, spreadsheet){

  // キャッシュを有効化（開発時はfalseに設定可能）
  const USE_CACHE = true;

  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  rememberContinuationTarget(ss, cacheOnly);
  const settings = getSettings(ss);
  Logger.log('適用設定: ' + JSON.stringify(settings)
    + ' / 株価上限の目安: ' + getMaxStockPrice(settings) + '円');
  
  showProgress('開始しています', 0);

  const iList = create_industrycodelist(USE_CACHE)

  if (iList == null || iList.length === 0){
    return finishProgress("業種リスト取得エラー: データを取得できませんでした");
  }

  let stockTable = []
  const totalIndustries = iList.length;
  let processedIndustries = 0;
  let cacheMissPages = 0;

  for ( let industry of iList){

      for ( let pageNo = 1; pageNo < CONFIG.MAX_PAGES; pageNo++ ){

        checkBudget();
        const result = cacheOnly
          ? getCachedStocksByPage(industry.CODE, pageNo, settings)
          : get_stocks_by_page(industry.CODE, pageNo, USE_CACHE, settings)

        if (result.cacheMiss) cacheMissPages++;
        if ( result.done ) break  // そのページにデータなし = 最終ページ超え

        result.stocks.forEach((stock) => {
          stock.INDUSTRY = industry.NAME;
          stock.IS_FINANCIAL = FINANCIAL_INDUSTRY_CODES.indexOf(industry.CODE) !== -1;
        });
        stockTable = stockTable.concat(result.stocks)
      }

      // 進捗を更新
      processedIndustries++;
      showProgress(stageMessage(1, '業種一覧を取得中', processedIndustries, totalIndustries, '業種'),
        stagePercent(0, 40, processedIndustries, totalIndustries));
  }

  Logger.log('取得した銘柄数: ' + stockTable.length
    + (cacheOnly ? '（キャッシュのみ、ミス ' + cacheMissPages + 'ページ）' : ''));

  if (stockTable.length === 0) {
    return finishProgress(cacheOnly
      ? "キャッシュにデータがありません。「銘柄データ更新」を先に実行してください（キャッシュは最大1時間で消えます）"
      : "条件に合う銘柄が見つかりませんでした");
  }
  if (cacheOnly && cacheMissPages > 0) {
    Logger.log('注意: キャッシュ切れのページが ' + cacheMissPages + ' 件あり、結果が一部欠けている可能性があります');
  }

  // 第1.5段階: 25日平均出来高による流動性フィルタ（候補だけ個別取得する）
  const beforeVolume = stockTable.length;
  checkBudget();
  stockTable = applyVolumeFilter(stockTable, settings, USE_CACHE, (done, total) => {
    showProgress(stageMessage(2, '出来高を確認中', done, total, '銘柄'), stagePercent(40, 80, done, total));
    checkBudget();
  }, cacheOnly);
  Logger.log('出来高フィルタ: ' + beforeVolume + '件 → ' + stockTable.length + '件');

  if (stockTable.length === 0) {
    return finishProgress("出来高条件に合う銘柄が見つかりませんでした");
  }

  // 配当の持続性チェック（候補が一番少ない時点で財務ページを取得。落とさず注意に書く）
  applyDividendSafety(stockTable, settings, USE_CACHE, (done, total) => {
    showProgress(stageMessage(3, '配当を確認中', done, total, '銘柄'), stagePercent(80, 95, done, total));
    checkBudget();
  }, cacheOnly);
  Logger.log('ここまでの経過時間: ' + Math.round((Date.now() - startedAt) / 1000) + '秒');

  // 特別利益で最終益が膨らんだ銘柄は、補正PERでミックス係数を出し直して閾値を判定し直す
  const beforeQuality = stockTable.length;
  stockTable = applyEarningsQuality(stockTable, settings);
  Logger.log('補正ミックス係数: ' + beforeQuality + '件 → ' + stockTable.length + '件');
  if (stockTable.length === 0) {
    return finishProgress("補正後のミックス係数で条件に合う銘柄が見つかりませんでした");
  }

  showProgress('pickup シートに書き込み中', 95);

  // 割安な順（補正後のミックス係数の小さい順）。同点はコード順（英字付きコードでも崩れないよう文字列比較）
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
  let header = ['コード','名称','業種','株価','前日比','騰落率','PER','補正PER','PBR','配当利回','ミックス係数','購入金額（単元）','出来高（25日平均）','予想配当性向','注意']
  let allData = [header]

  for( let stock of stockTable ){
    const rowData = [stock.TICKER, stock.NAME, stock.INDUSTRY, stock.PRICE, stock.CHANGE, stock.CHANGE_RATE, stock.PER,
      formatOrUnknown(stock.ADJUSTED_PER), stock.PBR, stock.DPS,
      stock.MIX, Number(stock.PRICE) * settings.unitShares, formatOrUnknown(stock.AVG_VOLUME),
      formatOrUnknown(stock.PAYOUT_RATIO), (stock.NOTES || []).join(' / ')]
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
      
      // 前日比と騰落率を右詰めに設定
      sheet.getRange(1, header.indexOf('前日比') + 1, allData.length, 2).setHorizontalAlignment('right');

      // シートに残った書式（通貨など）でミックス係数が「¥1」と丸まって見えたため、数値列の書式を毎回決める
      const numberFormats = { '補正PER': '0.00', 'ミックス係数': '0.00', '購入金額（単元）': '#,##0', '出来高（25日平均）': '#,##0', '予想配当性向': '0.0' };
      Object.keys(numberFormats).forEach((name) => {
        sheet.getRange(2, header.indexOf(name) + 1, stockTable.length, 1).setNumberFormat(numberFormats[name]);
      });
      
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
    finishProgress(errorMsg);
    throw e;
  }
  
  // 最終更新日時はヘッダーの右に1列空けて置く
  const lastUpdateCell = sheet.getRange(1, header.length + 2);
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  lastUpdateCell.setValue('最終更新: ' + formattedDate);
  lastUpdateCell.setFontSize(9);
  lastUpdateCell.setFontColor('#666666');

  // 閾値の試行（キャッシュで再判定）は混ぜず、通常実行の候補だけを記録する
  if (!cacheOnly) {
    try {
      appendHistory(ss, stockTable, settings, formattedDate);
    } catch (e) {
      Logger.log('history 追記エラー: ' + e.toString());
    }
  }
  
  const flagged = stockTable.filter((stock) => (stock.NOTES || []).length > 0).length;
  return finishProgress('候補 ' + stockTable.length + '件（うち注意あり ' + flagged + '件）を pickup に出力しました');
}

/** 既存の history と列がずれないよう、新しい列は末尾に足す。ミックス係数は補正PERベース */
const HISTORY_HEADER = ['実行日時','コード','名称','業種','株価','PER','PBR','配当利回','ミックス係数','出来高（25日平均）','予想配当性向','注意','適用設定','補正PER'];

/**
 * 候補を history シートの末尾に追記する（後で成績を比べる効果測定用）。
 * 適用設定を JSON で残し、どの閾値で選んだ候補かを後から区別できるようにする。
 */
function appendHistory(ss, stocks, settings, runAt) {
  if (!stocks || stocks.length === 0) return;
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HISTORY_HEADER.length).setValues([HISTORY_HEADER]);
    sheet.setFrozenRows(1);
  } else {
    // 列を足す前に作られたシートは見出しが足りないので書き足す
    const headerRange = sheet.getRange(1, 1, 1, HISTORY_HEADER.length);
    const current = headerRange.getValues()[0];
    if (HISTORY_HEADER.some((name, i) => current[i] !== name)) {
      headerRange.setValues([HISTORY_HEADER]);
    }
  }
  const settingsJson = JSON.stringify(settings);
  const rows = stocks.map((stock) => [runAt, stock.TICKER, stock.NAME, stock.INDUSTRY, stock.PRICE,
    stock.PER, stock.PBR, stock.DPS, stock.MIX, formatOrUnknown(stock.AVG_VOLUME),
    formatOrUnknown(stock.PAYOUT_RATIO), (stock.NOTES || []).join(' / '), settingsJson,
    formatOrUnknown(stock.ADJUSTED_PER)]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HISTORY_HEADER.length).setValues(rows);
}

/** 取得できなかった値（null/undefined）をシート上で「不明」と表示する */
function formatOrUnknown(value) {
  return (value === null || value === undefined) ? '不明' : value;
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

const PROGRESS_STAGE_COUNT = 3;

/** 例: "【2/3】出来高を確認中: 37 / 493 銘柄" */
function stageMessage(stage, label, done, total, unit) {
  return '【' + stage + '/' + PROGRESS_STAGE_COUNT + '】' + label + ': '
    + done.toLocaleString() + ' / ' + total.toLocaleString() + ' ' + unit;
}

/** 段階ごとに割り当てたバーの範囲 [from, to] の中で進める */
function stagePercent(from, to, done, total) {
  if (!total) return to;
  return Math.round(from + (to - from) * done / total);
}

const PROGRESS_CACHE_KEY = 'pickup_progress';
const PROGRESS_CACHE_DURATION = 600;

/**
 * 進捗をユーザー単位のキャッシュに書く。サイドバーが getProgress で読みに来る。
 * 進捗の書き込み失敗で本処理を止めないよう、例外は握りつぶしてログだけ残す。
 */
function showProgress(message, percent) {
  writeProgress({ message: message, percent: percent, done: false });
}

/** 終了を記録し、同じメッセージを返す（runPickup の戻り値にそのまま使う） */
function finishProgress(message) {
  Logger.log(message);
  writeProgress({ message: message, percent: 100, done: true });
  return message;
}

function writeProgress(state) {
  try {
    CacheService.getUserCache().put(PROGRESS_CACHE_KEY, JSON.stringify(state), PROGRESS_CACHE_DURATION);
  } catch (e) {
    Logger.log('進捗の書き込みエラー: ' + e.toString());
  }
}

/** サイドバーからポーリングで呼ばれる */
function getProgress() {
  const cached = CacheService.getUserCache().get(PROGRESS_CACHE_KEY);
  return cached ? JSON.parse(cached) : null;
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
  const cacheKey = getCacheKey(STOCKS_CACHE_TYPE, encodeURIComponent(industryCode), pageNo);

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
  const cacheKey = getCacheKey(STOCKS_CACHE_TYPE, encodeURIComponent(industryCode), pageNo);
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

  // PER（"倍"と桁区切りを除去、"ー"は0扱い）。parseFloat("1,052.5") は 1 になるため、カンマは必ず先に消す
  const perText = removeHtmlTags(tdMatches[2]).replace(/[倍,]/g, '').trim();
  res.PER = (perText === 'ー' || perText === '') ? 0 : parseFloat(perText);

  // PBR
  const pbrText = removeHtmlTags(tdMatches[3]).replace(/[倍,]/g, '').trim();
  res.PBR = (pbrText === 'ー' || pbrText === '') ? 0 : parseFloat(pbrText);

  // 配当利回り（"%"を除去）
  const dpsText = removeHtmlTags(tdMatches[4]).replace(/[%,]/g, '').trim();
  res.DPS = (dpsText === 'ー' || dpsText === '') ? 0 : parseFloat(dpsText);

  // グレアムのミックス係数（PER・PBRのどちらかが欠損した場合は判定不能として null）
  res.MIX = (res.PER > 0 && res.PBR > 0) ? Math.round(res.PER * res.PBR * 1000) / 1000 : null;

  return res;
}