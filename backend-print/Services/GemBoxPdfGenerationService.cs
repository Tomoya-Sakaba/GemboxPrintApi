using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using GemBox.Spreadsheet;

namespace backend_print.Services
{
    /// <summary>
    /// GemBoxでExcel→PDFを生成（backend-print 側に隔離）
    /// </summary>
    public class GemBoxPdfGenerationService
    {
        private readonly string _tempPath;

        public GemBoxPdfGenerationService()
        {
            // OSのテンポラリフォルダ。
            // テンプレExcelをコピーした「作業用xlsx」と、生成した「作業用pdf」をここに置く。
            _tempPath = Path.GetTempPath();

            // GemBoxライセンスキー（Web.config）。
            // 未設定の場合は FREE-LIMITED-KEY で動作（機能/ページ制限がかかる可能性あり）。
            var key = ConfigurationManager.AppSettings["GemBoxSpreadsheetLicenseKey"];
            SpreadsheetInfo.SetLicense(string.IsNullOrWhiteSpace(key) ? "FREE-LIMITED-KEY" : key);
        }

        public Stream GeneratePdf(string templatePath, Dictionary<string, object> data)
        {
            // ここは「テンプレのコピー → 埋め込み → PDF変換 → MemoryStreamで返す」までを担当する。
            // API側（Controller）はこのStreamをそのままHTTPレスポンスに載せる。

            // 作業用ファイルパス（finallyで消す）
            string tempExcelPath = null;
            string tempPdfPath = null;

            // 速度計測（ログ用。現在はログ呼び出しをコメントアウトしている）
            var sw = Stopwatch.StartNew();

            try
            {
                // --- 1) テンプレExcelを作業用にコピー ---
                // 元テンプレを直接編集すると、同時実行時に競合する/テンプレが壊れる可能性があるため、
                // 必ず一時ファイルにコピーしてから編集する。
                tempExcelPath = Path.Combine(_tempPath, $"gembox_{Guid.NewGuid()}.xlsx");
                // SimpleFileLogger.Log(GetLogPath(), $"GeneratePdf start. templatePath='{templatePath}', tempExcel='{tempExcelPath}'");
                File.Copy(templatePath, tempExcelPath, true);

                // --- 2) プレースホルダ置換（Excel編集） ---
                // tempExcelPath の中の {{...}} を data で置換する。
                // SimpleFileLogger.Log(GetLogPath(), $"EmbedData start. elapsedMs={sw.ElapsedMilliseconds}");
                EmbedData(tempExcelPath, data);
                // SimpleFileLogger.Log(GetLogPath(), $"EmbedData end. elapsedMs={sw.ElapsedMilliseconds}");

                // --- 3) Excel → PDF 変換 ---
                // GemBoxは Save(pdfPath) でPDF出力できる。
                tempPdfPath = Path.Combine(_tempPath, $"gembox_{Guid.NewGuid()}.pdf");
                // SimpleFileLogger.Log(GetLogPath(), $"ConvertExcelToPdf start. tempPdf='{tempPdfPath}', elapsedMs={sw.ElapsedMilliseconds}");
                ConvertExcelToPdf(tempExcelPath, tempPdfPath);
                // SimpleFileLogger.Log(GetLogPath(), $"ConvertExcelToPdf end. pdfBytes={new FileInfo(tempPdfPath).Length}, elapsedMs={sw.ElapsedMilliseconds}");

                // --- 4) PDFファイルをメモリへ読み込み、Streamで返す ---
                // API側でレスポンスに載せやすいよう MemoryStream にする。
                var pdfStream = new MemoryStream(File.ReadAllBytes(tempPdfPath));
                // 読み込み直後はPositionが末尾になり得るので、明示的に先頭へ戻す。
                pdfStream.Position = 0;
                // SimpleFileLogger.Log(GetLogPath(), $"GeneratePdf done. elapsedMs={sw.ElapsedMilliseconds}");
                return pdfStream;
            }
            catch (Exception ex)
            {
                // 例外は上位（Controller）で 500 として返す。
                // ここでは追加処理はせず、必要ならログを有効化して原因調査する。
                // SimpleFileLogger.Log(GetLogPath(), $"GeneratePdf ERROR. elapsedMs={sw.ElapsedMilliseconds}. {ex}");
                throw;
            }
            finally
            {
                // --- 5) 作業ファイルの後始末 ---
                // 例外が出ても temp ファイルが溜まらないよう、必ず削除を試みる。
                Cleanup(tempExcelPath);
                Cleanup(tempPdfPath);
            }
        }

        private string GetLogPath()
        {
            // ログファイルパス（Web.config）
            return ConfigurationManager.AppSettings["GemBoxLogFilePath"];
        }

        private void EmbedData(string excelPath, Dictionary<string, object> data)
        {
            // --- Excelの埋め込み処理 ---
            // 1) Excelロード
            // 2) 明細（table）展開
            // 3) 単票プレースホルダ置換
            // 4) 保存

            // 1) Excelをロード（.xlsx）
            var workbook = ExcelFile.Load(excelPath);

            // 今回はテンプレ1枚運用のため先頭シートを対象にする。
            // シートが複数ある帳票に拡張する場合は Worksheets をループする。
            var ws = workbook.Worksheets[0];

            // 2) 明細（テーブル）を展開する。
            // {{table:history}} のようなマーカー行を見つけたら、その行を複製して行数ぶん埋める。
            ExpandTableRegions(ws, data);

            // 3) 単票プレースホルダ置換のための正規表現。
            // 例: "機器コード: {{equipment_code}}" の {{equipment_code}} 部分を検出する。
            var regex = new Regex(@"\{\{(.+?)\}\}");

            // UsedRange（使用範囲）を取得。
            // ws.Cells 全走査は膨大で遅い/ハングに見えることがあるため、必ず使用範囲だけ走査する。
            var used = ws.GetUsedCellRange(true);
            if (used == null)
            {
                // 置換対象が見つからなくても、編集結果（明細展開など）があるかもしれないので保存はする。
                workbook.Save(excelPath);
                return;
            }

            // 行・列を UsedRange の範囲で走査する。
            for (int r = used.FirstRowIndex; r <= used.LastRowIndex; r++)
            {
                for (int c = used.FirstColumnIndex; c <= used.LastColumnIndex; c++)
                {
                    // 対象セル
                    var cell = ws.Cells[r, c];

                    // 文字列セルのみ置換対象（数値/日付/数式などは触らない）
                    if (cell.ValueType != CellValueType.String) continue;

                    // セル文字列
                    var s = cell.StringValue;

                    // 空・空白のみは対象外
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    // "{{" が無いセルは対象外（正規表現の無駄打ち回避）
                    if (s.IndexOf("{{", StringComparison.Ordinal) < 0) continue;

                    // セル内の {{key}} を data[key] に置換する。
                    // 見つからないキーは空文字にする（テンプレ側の書き間違いでも処理は継続）
                    var replaced = regex.Replace(s, m =>
                    {
                        // {{ ... }} の中身（前後空白は除去）
                        var key = m.Groups[1].Value.Trim();

                        // data にキーがあれば文字列化して返す
                        if (data.TryGetValue(key, out var v))
                            return FormatValue(v);

                        // 無い場合は空文字（置換）
                        return "";
                    });

                    // 変化があったときだけ書き戻す（無駄な変更を減らす）
                    if (replaced != s) cell.Value = replaced;
                }
            }

            // 4) 置換結果を同じパスに保存
            workbook.Save(excelPath);
        }

        private void ExpandTableRegions(ExcelWorksheet ws, Dictionary<string, object> data)
        {
            // --- 明細（テーブル）展開 ---
            // テンプレの“行テンプレ”に以下を配置する:
            // - {{table:history}}  ← テーブル開始マーカー（どのセルでもOK）
            // - 同じ行に {{history.date}} / {{history.note}} のようなセルを配置
            //
            // data 側は:
            // - data["history"] = IEnumerable<Dictionary<string, object>>
            //
            // 展開処理:
            // - マーカー行を見つけたら、その行を必要行数ぶん InsertCopy で複製
            // - 複製した各行について FillTableRow で {{history.xxx}} を埋める

            var tableStartRegex = new Regex(@"\{\{\s*table\s*:\s*([a-zA-Z0-9_]+)\s*\}\}");
            var used = ws.GetUsedCellRange(true);
            if (used == null) return;

            for (int r = used.FirstRowIndex; r <= used.LastRowIndex; r++)
            {
                for (int c = used.FirstColumnIndex; c <= used.LastColumnIndex; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell.ValueType != CellValueType.String) continue;
                    var s = cell.StringValue;
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    // このセルが {{table:xxx}} を含むか判定
                    var m = tableStartRegex.Match(s);
                    if (!m.Success) continue;

                    // tableKey = "history" 等
                    var tableKey = m.Groups[1].Value.Trim();

                    // マーカー文字列は出力不要なので消す（セルを空にする）
                    cell.Value = "";

                    // data に tableKey が無い場合は何もしない（テンプレ側の書き間違い想定）
                    if (!data.TryGetValue(tableKey, out var rowsObj)) return;

                    // rowsObj を IEnumerable<Dictionary<string, object>> として扱う
                    var rows = rowsObj as IEnumerable<Dictionary<string, object>>;
                    if (rows == null) return;

                    // 複数回数えないようにList化
                    var list = rows.ToList();

                    if (list.Count == 0)
                    {
                        // 明細が0件の場合:
                        // - 行は残す（レイアウト維持）
                        // - {{history.xxx}} だけ消して空欄行にする
                        ClearTableRowPlaceholders(ws, r, tableKey);
                        return;
                    }

                    if (list.Count > 1)
                    {
                        // 明細が2件以上の場合:
                        // - 現在行(r)がテンプレ行なので、下に (件数-1) 行ぶんコピー挿入する
                        // - InsertCopy は “書式/罫線/結合” を含めてコピーされる
                        ws.Rows.InsertCopy(r + 1, list.Count - 1, ws.Rows[r]);
                    }

                    // 各行に対して {{history.xxx}} を埋める
                    for (int i = 0; i < list.Count; i++)
                        FillTableRow(ws, r + i, tableKey, list[i]);

                    // 1つのテーブルを展開したら終了（複数テーブル対応は必要になったら拡張）
                    return;
                }
            }
        }

        private void FillTableRow(ExcelWorksheet ws, int rowIndex, string tableKey, Dictionary<string, object> rowData)
        {
            // rowIndex の行にある {{history.xxx}} を rowData["xxx"] で置換する。
            // tableKey は "history" など。
            var regex = new Regex(@"\{\{\s*" + Regex.Escape(tableKey) + @"\.([a-zA-Z0-9_]+)\s*\}\}");
            var used = ws.GetUsedCellRange(true);
            if (used == null) return;

            for (int c = used.FirstColumnIndex; c <= used.LastColumnIndex; c++)
            {
                var cell = ws.Cells[rowIndex, c];
                if (cell.ValueType != CellValueType.String) continue;
                var s = cell.StringValue;
                if (string.IsNullOrWhiteSpace(s)) continue;

                // {{history.col}} を rowData[col] に置換
                var replaced = regex.Replace(s, m =>
                {
                    // col（= history の列名部分）
                    var key = m.Groups[1].Value.Trim();
                    if (rowData != null && rowData.TryGetValue(key, out var v))
                        return FormatValue(v);
                    return "";
                });

                if (replaced != s) cell.Value = replaced;
            }
        }

        private void ClearTableRowPlaceholders(ExcelWorksheet ws, int rowIndex, string tableKey)
        {
            // 明細0件時に、行テンプレ内の {{history.xxx}} だけを空文字にする。
            var regex = new Regex(@"\{\{\s*" + Regex.Escape(tableKey) + @"\.[a-zA-Z0-9_]+\s*\}\}");
            var used = ws.GetUsedCellRange(true);
            if (used == null) return;

            for (int c = used.FirstColumnIndex; c <= used.LastColumnIndex; c++)
            {
                var cell = ws.Cells[rowIndex, c];
                if (cell.ValueType != CellValueType.String) continue;
                var s = cell.StringValue;
                if (string.IsNullOrWhiteSpace(s)) continue;
                var replaced = regex.Replace(s, "");
                if (replaced != s) cell.Value = replaced;
            }
        }

        private string FormatValue(object value)
        {
            // Excelセルに埋め込む際の文字列表現。
            // ここを拡張すると、数値のカンマや日時フォーマットなども統一できる。
            if (value == null) return "";
            if (value is DateTime dt) return dt.ToString("yyyy/MM/dd");
            return value.ToString();
        }

        private void ConvertExcelToPdf(string excelPath, string pdfPath)
        {
            // GemBoxによる変換:
            // - ExcelFile.Load でxlsxを読み
            // - Save(pdfPath) でPDFとして書き出す
            // ※テンプレの印刷範囲やセル結合/色/罫線などは、基本的にExcel側の設定を反映する。
            var workbook = ExcelFile.Load(excelPath);
            workbook.Save(pdfPath);
        }

        private void Cleanup(string path)
        {
            // 一時ファイルの削除（失敗しても業務処理は止めない）
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }
    }
}

