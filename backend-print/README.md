## 概要

`backend-print` は **Excelテンプレート（.xlsx）に値を埋め込み、GemBox.SpreadsheetでPDF化して返す** Web API です。
DBアクセスは行わず、呼び出し側（例: 別の backend）が **テンプレ名 + data/tables/pictures** を組み立てて POST します。

## エンドポイント

- `POST /api/print/gembox/pdf`

## リクエスト（JSON）

### ボディ型

`backend_print.Models.DTOs.GemBoxPrintRequestDto`

- **templateFileName**（必須）: `string`
- **data**（任意）: `Record<string, any>`（単票）
- **tables**（任意）: `Record<string, Array<Record<string, any>>>`（明細）
- **pictures**（任意）: `Record<string, string>`（画像）

### 必須条件（これを満たさないと失敗します）

- **`templateFileName` が必須**
  - ファイル名のみ許可（パス混入不可）
  - 拡張子は **`.xlsx` 固定**
- **`data` / `tables` / `pictures` のいずれかに値が必要**
  - 3つとも空（または未指定）だと 400 になります

### 送信例（curl）

```bash
curl -X POST "http://localhost:62169/api/print/gembox/pdf" ^
  -H "Content-Type: application/json" ^
  -H "Accept: application/pdf" ^
  -H "X-Correlation-Id: demo-001" ^
  --data "{ \"templateFileName\": \"demo_gembox.xlsx\", \"data\": { \"title\": \"GemBox demo\" } }" ^
  --output out.pdf
```

## テンプレート（Excel）側の書き方

### 単票（data）

Excelセルに `{{key}}` を置きます。

- `data["title"] = "..."` → `{{title}}` が置換されます

### 明細（tables）

同一行に以下を並べる運用です。

- 行テンプレの開始マーカー: `{{table:items}}`
- 同じ行に列プレースホルダ: `{{items.name}}`, `{{items.qty}}` など

`tables["items"]` が 0件なら展開されず、1件以上なら行が増殖します（テンプレ上の配置を基準に下方向へ展開）。

### 画像（pictures）

画像は **セル全体が `{{key}}` のときだけ**埋め込み対象になります（文章中の一部に画像を混ぜる用途は想定しません）。

- `pictures["pic_1_1"] = "C:\\app_data\\picuture\\test1.png"` のように **絶対パス**を渡せます
- `pictures["pic_1_1"] = "test1.png"` のように **ファイル名のみ**を渡すこともできます  
  - この場合は `Web.config` の `GemBoxPictureBasePath` と結合して探します

対応拡張子: `.png .jpg .jpeg .gif .bmp .tif .tiff .svg .emf .wmf`

## レスポンス

- 成功時: `200 OK`
  - `Content-Type: application/pdf`
  - ボディ: PDFバイナリ

※ このAPIは **`Content-Disposition`（filename）を付けません**。保存名は呼び出し側（フロント等）で決める想定です。

## エラー（主なもの）

- `400 BadRequest`
  - ボディが空
  - `templateFileName` が未指定/不正（パス混入、拡張子が .xlsx 以外など）
  - `data` / `tables` / `pictures` が全部空
- `404 NotFound`
  - テンプレートファイルが存在しない

## 複数シートについて

- テンプレに複数シートがある場合、**全シートに対して置換処理**を行います
- PDF化は `SelectionType = EntireFile` のため、**全シートを1つのPDF**として出力します

## 設定（Web.config / appSettings）

- **`BReportTemplateBasePath`**
  - テンプレ `.xlsx` を置くフォルダ
- **`GemBoxSpreadsheetLicenseKey`**
  - 空だと無料版（制限あり）
- **`GemBoxPictureBasePath`**
  - pictures の値が相対（例: `test1.png`）のときのベースフォルダ

## 設定のDB管理（m_key）

パス等を DB で管理する運用の場合、`dbo.m_key`（`k`/`v`）に設定値を入れておくと
backend-print は **DBからのみ取得**します（m_key に無い場合はエラー）。

- 例（DB側）:
  - `k = 'BReportTemplateBasePath'` → `v = '~/App_Data/b-templates'`
  - `k = 'GemBoxPictureBasePath'` → `v = 'C:\app_data\picuture'`

※ DBから読むために、backend-print の `Web.config` に `connectionStrings` の `MyDbConnection` を設定してください。

