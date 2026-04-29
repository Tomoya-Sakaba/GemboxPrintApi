using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace backend_print.Utils
{
    public static class PrintGemBoxRequestUtils
    {
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
        public static Dictionary<string, object> MergeToGemBoxData(Models.DTOs.GemBoxPrintRequestDto request)
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
        public static Dictionary<string, string> BuildPicturesDictionary(Models.DTOs.GemBoxPrintRequestDto request)
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

