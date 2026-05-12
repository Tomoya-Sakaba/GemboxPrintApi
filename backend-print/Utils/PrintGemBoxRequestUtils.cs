using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using backend_print.Models.DTOs;
using Newtonsoft.Json.Linq;

namespace backend_print.Utils
{
    /// <summary>
    /// <see cref="PrintGemBoxRequestUtils.TryPrepareGemBoxPrint"/> が成功したときのテンプレパスとマージ済みデータ。
    /// </summary>
    public sealed class GemBoxPrintPreparedContext
    {
        public string TemplatePath { get; set; }
        public string TemplateFileName { get; set; }
        public Dictionary<string, object> MergedData { get; set; }
        public Dictionary<string, string> Pictures { get; set; }
    }

    /// <summary>
    /// GemBox PDF / Excel 共通のリクエスト検証結果。
    /// </summary>
    public readonly struct GemBoxPrintPrepareResult
    {
        public bool Ok { get; }
        public GemBoxPrintPreparedContext Prepared { get; }
        public HttpStatusCode ErrorStatus { get; }
        public string ErrorMessage { get; }
        /// <summary>ログ用（correlationId は呼び出し側で付与）。</summary>
        public string FailureLogDetail { get; }

        private GemBoxPrintPrepareResult(
            bool ok,
            GemBoxPrintPreparedContext prepared,
            HttpStatusCode errorStatus,
            string errorMessage,
            string failureLogDetail)
        {
            Ok = ok;
            Prepared = prepared;
            ErrorStatus = errorStatus;
            ErrorMessage = errorMessage;
            FailureLogDetail = failureLogDetail;
        }

        public static GemBoxPrintPrepareResult Success(GemBoxPrintPreparedContext ctx) =>
            new GemBoxPrintPrepareResult(true, ctx, default, null, null);

        public static GemBoxPrintPrepareResult Failure(HttpStatusCode status, string message, string failureLogDetail) =>
            new GemBoxPrintPrepareResult(false, null, status, message, failureLogDetail);
    }

    public static class PrintGemBoxRequestUtils
    {
        /// <summary>
        /// PDF / Excel 共通: ボディ・テンプレ名・ファイル存在・data/tables/pictures 非空を検証し、埋め込み用データを組み立てる。
        /// </summary>
        public static GemBoxPrintPrepareResult TryPrepareGemBoxPrint(
            GemBoxPrintRequestDto request,
            string templateBasePath)
        {
            if (request == null)
            {
                return GemBoxPrintPrepareResult.Failure(
                    HttpStatusCode.BadRequest,
                    "リクエストボディが空です。",
                    "リクエスト検証失敗（ボディが空）");
            }

            var templateFile = (request.TemplateFileName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(templateFile) ||
                !IsSafeFileNameWithExtension(templateFile, ".xlsx"))
            {
                return GemBoxPrintPrepareResult.Failure(
                    HttpStatusCode.BadRequest,
                    "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。",
                    $"リクエスト検証失敗（templateFileNameが不正） templateFileName='{request.TemplateFileName}'");
            }

            var templatePath = Path.Combine(templateBasePath, templateFile);
            if (!File.Exists(templatePath))
            {
                return GemBoxPrintPrepareResult.Failure(
                    HttpStatusCode.NotFound,
                    "テンプレートファイルが見つかりません。",
                    $"リクエスト検証失敗（テンプレート未存在） templatePath='{templatePath}'");
            }

            var merged = MergeToGemBoxData(request);
            var picturesMap = BuildPicturesDictionary(request);
            if (merged.Count == 0 && picturesMap.Count == 0)
            {
                return GemBoxPrintPrepareResult.Failure(
                    HttpStatusCode.BadRequest,
                    "印刷データが指定されていません。data / tables / pictures のいずれかに値を指定してください。",
                    "リクエスト検証失敗（data/tables/pictures が空）");
            }

            return GemBoxPrintPrepareResult.Success(new GemBoxPrintPreparedContext
            {
                TemplatePath = templatePath,
                TemplateFileName = templateFile,
                MergedData = merged,
                Pictures = picturesMap,
            });
        }

        /// <summary>
        /// <c>addPdfPath</c>（末尾結合用）。空なら成功で <paramref name="resolvedPath"/> は null。
        /// 値ありのときはフルパス・.pdf・ファイル存在を検証する。
        /// </summary>
        public static bool TryValidateAddPdfPathForAppend(
            string addPdfPathRaw,
            out string resolvedPath,
            out HttpStatusCode errorStatus,
            out string errorMessage,
            out string failureLogDetail)
        {
            resolvedPath = null;
            errorStatus = default;
            errorMessage = null;
            failureLogDetail = null;

            var t = (addPdfPathRaw ?? "").Trim();
            if (t.Length == 0)
                return true;

            if (!Path.IsPathRooted(t))
            {
                errorStatus = HttpStatusCode.BadRequest;
                errorMessage = "addPdfPath はローカル絶対パス（フルパス）で指定してください。";
                failureLogDetail = $"addPdfPath 検証失敗（フルパスではない） addPdfPath='{t}'";
                return false;
            }

            if (!t.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                errorStatus = HttpStatusCode.BadRequest;
                errorMessage = "addPdfPath は .pdf ファイルのフルパスを指定してください。";
                failureLogDetail = $"addPdfPath 検証失敗（拡張子） addPdfPath='{t}'";
                return false;
            }

            if (!File.Exists(t))
            {
                errorStatus = HttpStatusCode.NotFound;
                errorMessage = "addPdfPath で指定されたPDFが見つかりません。";
                failureLogDetail = $"addPdfPath 検証失敗（ファイルなし） path='{t}'";
                return false;
            }

            resolvedPath = t;
            return true;
        }

        public static string GetCorrelationId(HttpRequestMessage request)
        {
            try
            {
                if (request != null && request.Headers != null &&
                    request.Headers.TryGetValues("X-Correlation-Id", out var values))
                {
                    var v = values?.FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                }
            }
            catch
            {
            }
            return "-";
        }

        /// <summary>
        /// data（単票）と tables（明細）のみをマージする。画像は <see cref="BuildPicturesDictionary"/> で別途渡す。
        /// </summary>
        public static Dictionary<string, object> MergeToGemBoxData(GemBoxPrintRequestDto request)
        {
            var merged = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            if (request?.Data != null)
            {
                foreach (var kv in request.Data)
                    merged[kv.Key] = NormalizeValue(kv.Value);
            }

            if (request?.Tables != null)
            {
                foreach (var kv in request.Tables)
                {
                    var rows = kv.Value ?? new List<Dictionary<string, object>>();
                    var list = rows.Select(row =>
                    {
                        var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                        if (row == null) return d;
                        foreach (var c in row)
                            d[c.Key] = NormalizeValue(c.Value);
                        return d;
                    }).ToList();

                    merged[kv.Key] = list;
                }
            }

            return merged;
        }

        /// <summary>
        /// request.Pictures を GemBoxPdfGenerationService 用の辞書にする（dataとはマージしない）。
        /// </summary>
        public static Dictionary<string, string> BuildPicturesDictionary(GemBoxPrintRequestDto request)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (request?.Pictures == null) return d;
            foreach (var kv in request.Pictures)
            {
                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                var v = NormalizeValue(kv.Value);
                d[kv.Key.Trim()] = v?.ToString() ?? "";
            }
            return d;
        }

        public static object NormalizeValue(object v)
        {
            if (v == null) return "";
            if (v is JToken jt)
            {
                if (jt.Type == JTokenType.Null) return "";
                if (jt.Type == JTokenType.Date)
                {
                    var dt = jt.ToObject<DateTime>();
                    return dt;
                }
                return jt.Type == JTokenType.String ? jt.ToString() : jt.ToObject<object>();
            }
            return v;
        }

        /// <summary>
        /// パストラバーサル防止: ファイル名のみ。拡張子を指定して検証する。
        /// </summary>
        public static bool IsSafeFileNameWithExtension(string name, string extensionWithDot)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (string.IsNullOrWhiteSpace(extensionWithDot)) return false;
            if (!extensionWithDot.StartsWith(".", StringComparison.Ordinal)) return false;

            // ファイル名のみ許可（パス混入を拒否）
            var f = Path.GetFileName(name);
            if (!string.Equals(f, name, StringComparison.OrdinalIgnoreCase)) return false;

            // 不正文字を拒否
            if (f.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;

            // 拡張子チェック（.pdf / .xlsx）
            return f.EndsWith(extensionWithDot, StringComparison.OrdinalIgnoreCase);
        }
    }
}

