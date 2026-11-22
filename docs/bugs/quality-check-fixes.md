# 品質チェックで見つかった不具合5件（2026-11-22）

AI（musecode）による品質チェックで見つかり、Claude によるダブルチェックで実在を確認した不具合。回帰テストは `tests/quality_fixes.test.js`（`node --test tests/quality_fixes.test.js`）。

## 1. PER/PBR 欠損・負値の銘柄がスクリーニングを通過する（最重要）

- 症状: PER または PBR が `ー` の銘柄が、配当利回り 3.5% 以上なら結果に混ざる
- 原因: `get_stock_data` が `ー` を 0 に変換し、`MIX = 0` が `MIX < 11` を通過していた。負の PER/PBR でも `MIX` が負になって通過していた
- 対処: `PER > 0 && PBR > 0` のときだけ `MIX` を計算し、それ以外は `null` にする。フィルタ側でも `MIX !== null` を条件に加えた

## 2. 終了タグがないと空文字を株価として返す（finance.gs）

- 原因: `get_neatprice` で `indexOf` の結果 `-1` を未チェックのまま `substring(0, -1)` に渡し、`""` が返っていた
- 対処: 終了タグが見つからない場合と、抽出した価格が空の場合は `PriceEmpty` を返す

## 3. PageEmpty がログに残らない（finance.gs）

- 原因: `getprice_kabutan` の失敗判定が `FetchError` と `PriceEmpty` だけだった
- 対処: 判定に `PageEmpty` を追加した。戻り値がエラー文字列であることは他の失敗と同じで変わらない

## 4. キャッシュクリアが一部のページしか消さず、失敗時も成功と表示する

- 原因: クリアはページ1〜20だけだったが、取得側は最大98ページ分をキャッシュする。成功トーストが try/catch の外にあった
- 対処: 取得ループと同じ `p < CONFIG.MAX_PAGES` の範囲を削除し、成功トーストを try の中へ移した

## 5. 銘柄名に `"` が含まれると HYPERLINK 数式が壊れる、書き込みが遅い

- 対処: `createYahooFinanceFormula` で `"` を `""` にエスケープし、1セルずつの `setFormula` を `setFormulas` による一括書き込みに変えた
- 補足（推測）: 銘柄名は `<abbr title="...">` から取るため、`"` は通常 `&quot;` になっていて、このエスケープが効く場面はまれと考えられる

## 再発防止・運用上の注意

- キャッシュには「フィルタ後の結果」が最大1時間残る。判定ロジックや閾値を変えてデプロイした直後は、メニューの「キャッシュクリア」を実行する。根本対応は [deferred/screener-improvements.md](../deferred/screener-improvements.md) の1番
- GAS 実環境での確認（`setFormulas` での書き込み、約3,200キーのキャッシュ削除）は未実施
