# J-Net21 RSS 補助金モニター

J-Net21 の支援情報ヘッドライン RSS を平日 1 日 1 回確認し、対象地域に合致する当日分の新着の補助金・助成金・融資情報を Slack Incoming Webhook に投稿する Google Apps Script です。送信内容は Google Spreadsheet にログとして保存します。

作成の背景など : https://muhenkou.net/?p=10433

## できること

- J-Net21 RSS の平日日次チェック
- 地域フィルタリング
- 複数件ヒット時は 1 つの Slack メッセージにまとめて投稿
- 対象が 0 件のときは Slack 通知を行わず、実行ログのみ保存
- Spreadsheet への実行ログ・明細ログ・送信済み状態の保存
- `item_log` から公開用列だけを抜き出し、別ブックの `public_item_log` シートへ同期
- GAS の時間主導トリガー作成

デフォルトの対象地域は `東京都,神奈川県,埼玉県,千葉県` です。

## ファイル構成

- `Code.js` - Apps Script 本体
- `appsscript.json` - マニフェスト
- `.clasp.json.example` - clasp 接続設定の例

## セットアップ手順

1. Google Spreadsheet を 1 つ作成する
2. その Spreadsheet に紐づく Apps Script プロジェクトを作成する
3. このディレクトリの内容を `clasp` で push する
4. Apps Script の `Script Properties` に `SLACK_WEBHOOK_URL` を設定する
5. Apps Script で `initializeProject` を 1 回実行する

`initializeProject` を実行すると次を自動で行います。

- 必要シートの作成
- `config` シートの初期値投入
- `runDaily` の平日トリガー作成

## Spreadsheet シート構成

### `config`

| key | value | description |
| --- | --- | --- |
| RSS_URL | `https://j-net21.smrj.go.jp/snavi/support/support.xml` | J-Net21 の RSS URL |
| TARGET_REGIONS | `東京都,神奈川県,埼玉県,千葉県` | カンマ区切りの対象都道府県 |
| DAILY_TRIGGER_HOUR | `19` | 平日実行時刻 |

### `state`

通知済み・既知アイテムの状態を保持します。同じ RSS 記事を次回以降に再送しないために使います。

### `run_log`

日次実行ごとのサマリーログです。

### `item_log`

各記事ごとの判定結果ログです。新しい情報が上に来るように記録されます。

### `public_item_log`

`item_log` から次の列だけを抜き出して保持する外部公開用シートです。公開先は別の Spreadsheet とし、書き込み先は `Script Properties` の `PUBLIC_SPREADSHEET_ID` で指定します。Slack に載せる公開URLは `PUBLIC_SPREADSHEET_URL` で別管理します。
こちらも新しい情報が上に来る順で同期されます。

- `executed_at`
- `title`
- `link`
- `published_at`
- `matched_regions`

## Script Properties

Apps Script のプロジェクト設定から、次を設定してください。

- `SLACK_WEBHOOK_URL`: Slack Incoming Webhook URL

Apps Script を Spreadsheet に紐づけず単体プロジェクトとして動かす場合は、追加で次も設定してください。

- `SPREADSHEET_ID`: ログを書き込む Spreadsheet の ID
- `PUBLIC_SPREADSHEET_ID`: 公開用 `public_item_log` を書き込む別 Spreadsheet の ID
- `PUBLIC_SPREADSHEET_URL`: Slack通知の末尾に載せる公開用 URL

## Slack 通知仕様

- 新着が複数件ある場合は 1 メッセージに連番で並べて送信します
- 各通知にはRSSの概要文も含めます
- Slack通知の末尾には、公開用 Spreadsheet へのリンクを付与します
- 新着が 0 件の場合は Slack には送信せず、`run_log` と `item_log` にのみ結果を残します
- 通知文面では「当日分」として扱います

## 運用方法

- 地域を変更したい場合は `config` シートの `TARGET_REGIONS` を更新します
- 実行時刻を変更したい場合は `config` シートの `DAILY_TRIGGER_HOUR` を更新し、その後 `ensureDailyTrigger` を実行します
- 手動確認したい場合は `runDaily` を実行します。手動実行は曜日に関係なく動作します
- 自動トリガーは月曜から金曜のみ作成されます

## clasp 例

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json
# .clasp.json の scriptId を設定
clasp push
```

## 前提と補足

- 地域判定は RSS のタイトルと概要文に含まれる都道府県名をもとに行います
- `全国` を含む情報は、対象地域に関わらず通知対象として扱います
- J-Net21 側の RSS 構造が変わった場合は、パーサーの調整が必要になる可能性があります
