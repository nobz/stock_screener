# Git コミットメッセージのルール

CRITICAL: All Git commit messages MUST follow the rules below.

## ルール

- Conventional Commits のフォーマットに従うこと
- メッセージ本文（description）は必ず日本語で書くこと
- 英語の description は一切使用しないこと

## フォーマット

```
<type>: <日本語の説明>
```

## type 一覧

| type | 用途 |
|------|------|
| `feat` | 新機能の追加 |
| `fix` | バグ修正 |
| `refactor` | 機能変更を伴わないコードの整理 |
| `chore` | ビルド設定・ツール・依存関係の変更 |
| `docs` | ドキュメントのみの変更 |
| `style` | コードの動作に影響しない書式変更 |
| `perf` | パフォーマンス改善 |
| `test` | テストの追加・修正 |

## 例

良い例:
- `feat: 銘柄データ取得機能を追加`
- `fix: データ構造の問題を解決`
- `refactor: キャッシュ機能を整理して高速化`
- `chore: claspセットアップ、.clasp.json追加`
- `feat: Yahoo!ファイナンスへのリンクを追加`

悪い例:
- `feat: Add stock data fetching feature` (descriptionが英語なのでNG)
- `銘柄データ取得機能を追加` (typeがないのでNG)
- `fix: fixed bug` (descriptionが英語なのでNG)
