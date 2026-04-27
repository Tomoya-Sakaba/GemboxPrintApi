using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using GemBox.Spreadsheet;
using log4net;
using Newtonsoft.Json.Linq;

namespace backend_print.Services
{
    /// <summary>
    /// GemBoxでExcel→PDFを生成（backend-print 側に隔離）
    /// </summary>
    public class GemBoxPdfGenerationService
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(GemBoxPdfGenerationService));

        private readonly string _tempPath;

        public GemBoxPdfGenerationService()
        {
            // OSのテンポラリフォルダ。
            // テンプレExcelをコピーした「作業用xlsx」と、生成した「作業用pdf」をここに置く。
            _tempPath = Path.GetTempPath();

            // GemBoxライセンスキー（Web.config）。本番では製品キーを設定する。
            var key = ConfigurationManager.AppSettings["GemBoxSpreadsheetLicenseKey"];
            SpreadsheetInfo.SetLicense(string.IsNullOrWhiteSpace(key) ? "FREE-LIMITED-KEY" : key);
        }

        /// <summary>テンプレをコピーし、単票・明細は <paramref name="data"/>、画像プレースホルダは <paramref name="pictures"/> だけで解決して PDF 化する。</summary>
        /// <param name="data">単票・明細（テーブル）。request.Data と request.Tables のマージ。</param>
        /// <param name="pictures">画像用。セル全体が <c>{{key}}</c> のときのみ参照（data とは別辞書）。</param>
        public Stream GeneratePdf(
            string templatePath,
            Dictionary<string, object> data,
            IDictionary<string, string> pictures)
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
                Log.Info($"PDF生成開始. templatePath='{templatePath}', tempExcel='{tempExcelPath}'");
                File.Copy(templatePath, tempExcelPath, true);

                // --- 2) プレースホルダ置換（Excel編集） ---
                // tempExcelPath の中の {{...}} を data / pictures で置換する。
                Log.Debug($"Excel埋め込み開始. elapsedMs={sw.ElapsedMilliseconds}");
                EmbedData(tempExcelPath, data, pictures);
                Log.Debug($"Excel埋め込み完了. elapsedMs={sw.ElapsedMilliseconds}");

                // --- 3) Excel → PDF 変換 ---
                // GemBoxは Save(pdfPath) でPDF出力できる。
                tempPdfPath = Path.Combine(_tempPath, $"gembox_{Guid.NewGuid()}.pdf");
                Log.Debug($"PDF変換開始. tempPdf='{tempPdfPath}', elapsedMs={sw.ElapsedMilliseconds}");
                ConvertExcelToPdf(tempExcelPath, tempPdfPath);
                Log.Debug($"PDF変換完了. pdfBytes={new FileInfo(tempPdfPath).Length}, elapsedMs={sw.ElapsedMilliseconds}");

                // --- 4) PDFファイルをメモリへ読み込み、Streamで返す ---
                // API側でレスポンスに載せやすいよう MemoryStream にする。
                var pdfStream = new MemoryStream(File.ReadAllBytes(tempPdfPath));
                // 読み込み直後はPositionが末尾になり得るので、明示的に先頭へ戻す。
                pdfStream.Position = 0;
                Log.Info($"PDF生成完了. elapsedMs={sw.ElapsedMilliseconds}");
                return pdfStream;
            }
            catch (Exception ex)
            {
                Log.Error($"PDF生成失敗. elapsedMs={sw.ElapsedMilliseconds}", ex);
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

        /// <summary>
        /// Excelの埋め込み処理
        /// </summary>
        /// <param name="excelPath">Excelファイルパス</param>
        /// <param name="data">データ</param>
        /// <param name="pictures">画像</param>
        private void EmbedData(
            string excelPath,
            Dictionary<string, object> data,
            IDictionary<string, string> pictures)
        {
            // 1) Excelをロード（.xlsx）
            var workbook = ExcelFile.Load(excelPath);

            // 全シートに対して同じ埋め込み処理を行う（1枚目に単票・2枚目に横置きの明細など、テンプレ側で分担可能）。
            // 各シートの用紙方向・余白は Excel の「ページ設定」をテンプレに保存しておく（コードでは上書きしない）。
            var regex = new Regex(@"\{\{(.+?)\}\}");
            for (int si = 0; si < workbook.Worksheets.Count; si++)
            {
                var ws = workbook.Worksheets[si];

                // 2) 明細（テーブル）を展開する（同一シートに {{table:xxx}} を複数置ける。上から順に展開）。
                ExpandTableRegions(ws, data);

                // UsedRange（使用範囲）を取得。
                // ws.Cells 全走査は膨大で遅い/ハングに見えることがあるため、必ず使用範囲だけ走査する。
                var used = ws.GetUsedCellRange(true);
                if (used == null)
                    continue;

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

                        // 画像: セル全体が {{key}} の場合のみ対象にする（文章中へ画像を埋める用途は想定しない）
                        // 例: 結合セルの枠内に {{picture1}} を置き、data["picture1"]="test1.png" を渡す。
                        var m0 = regex.Match(s);
                        if (m0.Success && m0.Value == s.Trim())
                        {
                            var key0 = m0.Groups[1].Value.Trim();
                            if (pictures != null &&
                                pictures.TryGetValue(key0, out var imgRef) &&
                                !string.IsNullOrWhiteSpace(imgRef) &&
                                TryEmbedPicture(ws, cell, imgRef.Trim().Trim('"')))
                            {
                                cell.Value = "";
                                continue;
                            }

                            // セル全体が {{key}} の場合は、可能な限り「型のまま」セットして Excel の表示形式に任せる。
                            if (data.TryGetValue(key0, out var rawScalar))
                            {
                                cell.Value = CoerceToCellValue(rawScalar);
                                continue;
                            }
                        }

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
            }

            // 3) 置換結果を同じパスに保存（印刷設定はテンプレのまま）
            // 数式（SUMなど）の結果を、埋め込み後の値で更新してから保存する。
            workbook.Calculate();
            workbook.Save(excelPath);
        }

        private bool TryEmbedPicture(ExcelWorksheet ws, ExcelCell cell, string imageReference)
        {
            if (ws == null || cell == null) return false;
            if (string.IsNullOrWhiteSpace(imageReference)) return false;

            // 画像ファイル参照は「絶対パス」または「GemBoxPictureBasePath + ファイル名」を許可する。
            var basePath = DbKeyValueConfig.GetRequiredString("GemBoxPictureBasePath");
            var path = imageReference.Trim().Trim('"');
            if (!Path.IsPathRooted(path))
                path = Path.Combine(basePath, path);

            // 拡張子が画像っぽいもののみ対象
            var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            switch (ext)
            {
                case ".png":
                case ".jpg":
                case ".jpeg":
                case ".gif":
                case ".bmp":
                case ".tif":
                case ".tiff":
                case ".svg":
                case ".emf":
                case ".wmf":
                    break;
                default:
                    return false;
            }

            if (!File.Exists(path)) return false;

            // 結合セルなら結合範囲、そうでなければ単セルを枠として使う
            var range = cell.MergedRange;
            var firstRow = range != null ? range.FirstRowIndex : cell.Row.Index;
            var firstCol = range != null ? range.FirstColumnIndex : cell.Column.Index;
            var lastRow = range != null ? range.LastRowIndex : cell.Row.Index;
            var lastCol = range != null ? range.LastColumnIndex : cell.Column.Index;

            var topLeft = ws.Cells[firstRow, firstCol].Name;
            var bottomRight = ws.Cells[lastRow, lastCol].Name;

            var picture = ws.Pictures.Add(path, topLeft, bottomRight);
            // 画像のアスペクト比をセルに収めて余白で中央寄せする。
            picture.Position.Mode = PositioningMode.MoveAndSize;
            var boxLeft = picture.Position.Left;
            var boxTop = picture.Position.Top;
            var boxW = picture.Position.Width;
            var boxH = picture.Position.Height;
            if (boxW <= 0 || boxH <= 0)
                return true;

            // 罫線へのかぶり回避: 枠の内側に 1px 相当の余白を作る
            const double insetPx = 1.0;
            var insetLeft = boxLeft + insetPx;
            var insetTop = boxTop + insetPx;
            var insetW = Math.Max(0, boxW - (insetPx * 2.0));
            var insetH = Math.Max(0, boxH - (insetPx * 2.0));
            if (insetW <= 0 || insetH <= 0)
                return true;

            int pixW;
            int pixH;
            try
            {
                using (var img = Image.FromFile(path))
                {
                    pixW = img.Width;
                    pixH = img.Height;
                }
            }
            catch
            {
                // If pixel size unknown (e.g. some formats), keep default stretch.
                return true;
            }

            if (pixW <= 0 || pixH <= 0)
                return true;

            var imgAspect = (double)pixW / pixH;
            var boxAspect = insetW / insetH;
            double fittedW;
            double fittedH;
            if (imgAspect > boxAspect)
            {
                fittedW = insetW;
                fittedH = insetW / imgAspect;
            }
            else
            {
                fittedH = insetH;
                fittedW = insetH * imgAspect;
            }

            var padX = Math.Max(0, (insetW - fittedW) / 2.0);
            var padY = Math.Max(0, (insetH - fittedH) / 2.0);

            picture.Position.Mode = PositioningMode.FreeFloating;
            picture.Position.Left = insetLeft + padX;
            picture.Position.Top = insetTop + padY;
            picture.Position.Width = fittedW;
            picture.Position.Height = fittedH;
            return true;
        }

        private void ExpandTableRegions(ExcelWorksheet ws, Dictionary<string, object> data)
        {
            // --- 明細（テーブル）展開（複数テーブル可） ---
            // 各「行テンプレ」に {{table:xxx}} と同じ行に {{xxx.col}} を並べる。
            // data["xxx"] = IEnumerable<Dictionary<string, object>>
            //
            // テンプレ上の複数マーカーは上から順に処理する。先行テーブルで行挿入すると
            // 後続マーカーの行番号がずれるため、挿入行数を累積オフセットに加算する。

            // {{table:キー名}} の形式を検出する正規表現（キー名は英数字とアンダースコア）。
            var tableStartRegex = new Regex(@"\{\{\s*table\s*:\s*([a-zA-Z0-9_]+)\s*\}\}");

            // シート内の使用セル範囲のみ走査する（全セル走査は避ける）。
            var used = ws.GetUsedCellRange(true);
            if (used == null) return;

            // 使用範囲を走査し、{{table:xxx}} を含むセルをすべて列挙する。
            var markers = new List<TableMarker>();
            for (int r = used.FirstRowIndex; r <= used.LastRowIndex; r++)
            {
                for (int c = used.FirstColumnIndex; c <= used.LastColumnIndex; c++)
                {
                    var cell = ws.Cells[r, c];
                    // 文字列セルのみ（数値・数式セルは置換マーカーとして扱わない）。
                    if (cell.ValueType != CellValueType.String) continue;
                    var s = cell.StringValue;
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    var m = tableStartRegex.Match(s);
                    if (!m.Success) continue;

                    // 見つかった座標とテーブルキー（例: parts / linked）を保持する。
                    markers.Add(new TableMarker(r, c, m.Groups[1].Value.Trim()));
                }
            }

            // マーカーが無ければテーブル展開は行わない。
            if (markers.Count == 0) return;

            // 上から下へ（行が若い順、同じ行なら左の列が優先）で安定ソートする。
            markers.Sort((a, b) =>
                a.TemplateRow != b.TemplateRow
                    ? a.TemplateRow.CompareTo(b.TemplateRow)
                    : a.MarkerColumn.CompareTo(b.MarkerColumn));

            // 同一行に {{table:A}} と {{table:B}} が並ぶ場合は、最初に見つかった列だけを採用する（重複行を除外）。
            var seenTemplateRows = new HashSet<int>();
            var ordered = new List<TableMarker>();
            foreach (var mk in markers)
            {
                if (seenTemplateRows.Contains(mk.TemplateRow)) continue;
                seenTemplateRows.Add(mk.TemplateRow);
                ordered.Add(mk);
            }

            // テンプレ上の行番号は「固定」だが、上のテーブルで行挿入すると下のマーカーの実際の行番号がずれる。
            // rowOffset に、これまでに挿入した「増えた行数」を足し、ExpandOneTableRegion には実際の行番号を渡す。
            var rowOffset = 0;
            foreach (var mk in ordered)
            {
                var row = mk.TemplateRow + rowOffset;
                // ExpandOneTableRegion が返すのは「下に増えた行数」。次のマーカー用オフセットに加算する。
                rowOffset += ExpandOneTableRegion(ws, row, mk.MarkerColumn, mk.TableKey, data, tableStartRegex);
            }
        }

        /// <summary>テンプレ上の行・列（ファイル読み込み直後の座標）</summary>
        private struct TableMarker
        {
            public readonly int TemplateRow;
            public readonly int MarkerColumn;
            public readonly string TableKey;

            public TableMarker(int templateRow, int markerColumn, string tableKey)
            {
                TemplateRow = templateRow;
                MarkerColumn = markerColumn;
                TableKey = tableKey;
            }
        }

        /// <summary>
        /// 1 テーブル分を展開し、テンプレートに対して「下に増えた行数」（件数-1、0件は0）を返す。
        /// </summary>
        private int ExpandOneTableRegion(
            ExcelWorksheet ws,
            int row,
            int markerColumn,
            string tableKey,
            Dictionary<string, object> data,
            Regex tableStartRegex)
        {
            // マーカーが置かれているセル（通常は {{table:tableKey}} が入っているセル）。
            var cell = ws.Cells[row, markerColumn];
            if (cell.ValueType != CellValueType.String)
                return 0;

            var raw = cell.StringValue;
            // まだ {{table:...}} でない（オフセットずれで別行を指している等）なら何もしない。
            if (string.IsNullOrWhiteSpace(raw) || !tableStartRegex.IsMatch(raw))
                return 0;

            // マーカー文字列は消し、同じ行に {{tableKey.列名}} だけが残る想定で以降の置換に進む。
            cell.Value = "";

            // data にテーブルキーが無い場合は、行テンプレ内の {{tableKey.xxx}} を空にするだけ。
            if (!data.TryGetValue(tableKey, out var rowsObj))
            {
                ClearTableRowPlaceholders(ws, row, tableKey);
                return 0;
            }

            // 明細は IEnumerable<Dictionary<string, object>> のみ対応（JSON からの配列行）。
            var rows = rowsObj as IEnumerable<Dictionary<string, object>>;
            if (rows == null)
            {
                ClearTableRowPlaceholders(ws, row, tableKey);
                return 0;
            }

            var list = rows.ToList();

            // 0件なら行は増やさず、プレースホルダだけ除去する。
            if (list.Count == 0)
            {
                ClearTableRowPlaceholders(ws, row, tableKey);
                return 0;
            }

            // 2件目以降は「テンプレの1行」をコピーして行を挿入する（1件目は既存行をそのまま使う）。
            if (list.Count > 1)
                ws.Rows.InsertCopy(row + 1, list.Count - 1, ws.Rows[row]);

            // 各行について、{{tableKey.列名}} を行データで置換する。
            for (int i = 0; i < list.Count; i++)
                FillTableRow(ws, row + i, tableKey, list[i]);

            // 呼び出し元の rowOffset 用: テンプレより増えた行数（件数 1 なら 0）。
            return list.Count > 1 ? list.Count - 1 : 0;
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

                // セル全体が {{tableKey.col}} の場合は、可能な限り「型のまま」セットして Excel の表示形式に任せる。
                var m0 = regex.Match(s);
                if (m0.Success && m0.Value == s.Trim())
                {
                    var key0 = m0.Groups[1].Value.Trim();
                    if (rowData != null && rowData.TryGetValue(key0, out var rawCell))
                        cell.Value = CoerceToCellValue(rawCell);
                    else
                        cell.Value = "";
                    continue;
                }

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

        private object CoerceToCellValue(object value)
        {
            if (value == null) return "";

            // JSON 由来の値（JToken）を素の .NET 型へ
            if (value is JValue jv)
                value = jv.Value;
            else if (value is JToken jt)
                value = jt.ToString();

            // Excel の表示形式に任せたいので、数値や日時はできる限り型のまま返す。
            // 先頭ゼロ等を保持したい場合は、送信側で文字列にする。
            switch (value)
            {
                case DateTime _:
                case bool _:
                case byte _:
                case sbyte _:
                case short _:
                case ushort _:
                case int _:
                case uint _:
                case long _:
                case ulong _:
                case float _:
                case double _:
                case decimal _:
                    return value;
            }

            return value.ToString();
        }

        private void ConvertExcelToPdf(string excelPath, string pdfPath)
        {
            // GemBoxによる変換:
            // - ExcelFile.Load でxlsxを読み
            // - Save(pdfPath, PdfSaveOptions) でPDFとして書き出す
            //
            // 印刷オプション（向き・余白・改ページ）はテンプレの各シートのページ設定をそのまま使う。
            //
            // 重要: PdfSaveOptions.SelectionType の既定は ActiveSheet のため、
            // 「アクティブなシートだけ」が PDF に出る。複数シート（縦の1枚目＋横の2枚目など）を1本のPDFにまとめるには EntireFile が必要。
            var workbook = ExcelFile.Load(excelPath);
            // PDF出力前に数式結果を更新する（テンプレのキャッシュ値を出さないため）。
            workbook.Calculate();
            var pdfOptions = new PdfSaveOptions
            {
                SelectionType = SelectionType.EntireFile
            };
            workbook.Save(pdfPath, pdfOptions);
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

