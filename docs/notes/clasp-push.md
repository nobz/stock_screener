# clasp push の注意

- `rootDir` が `.` なので、`.claspignore` がないと `tests/` の `.js` まで GAS に送られる。テストファイルは先頭で `require` を呼んでおり、GAS では読み込み時にエラーになって全関数が動かなくなる。`.claspignore` は送るファイルを `appsscript.json`、`finance.gs`、`stock_screener.gs`、`progress.html`（進捗サイドバー）に限定している。GAS に送る `.gs` や `.html` を増やしたら、`.claspignore` にも追記する。`progress.html` が送られないと、メニューから実行したときに `createTemplateFromFile` が失敗する
- 送られるファイルは `clasp show-file-status` で確認できる
- `clasp push` は GAS 側のファイルを丸ごと置き換える。GAS のエディタで直接編集していた場合、その内容は消える。不安なときは一時フォルダに `clasp clone <scriptId>` して差分を見てから push する
- マニフェスト（`appsscript.json`）の上書き確認で止まらないよう、`clasp push --force` を使う
- 判定ロジックを変えて push したら、スプレッドシートのメニュー「キャッシュクリア」を実行する（理由は [bugs/quality-check-fixes.md](../bugs/quality-check-fixes.md)）
