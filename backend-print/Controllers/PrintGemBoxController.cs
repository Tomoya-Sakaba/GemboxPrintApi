using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using backend_print.Models.DTOs;
using backend_print.Services;
using log4net;
using Newtonsoft.Json.Linq;

namespace backend_print.Controllers
{
    /// <summary>
    /// GemBox: テンプレExcelへデータを埋め込みPDF化するのみ（DBアクセスなし）。
    /// POST /api/print/gembox/pdf
    /// </summary>
    [RoutePrefix("api/print/gembox")]
    public class PrintGemBoxController : ApiController
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(PrintGemBoxController));

        private readonly GemBoxPdfGenerationService _pdfService;
        private readonly string _templateBasePath;

        public PrintGemBoxController()
        {
            _pdfService = new GemBoxPdfGenerationService();
            _templateBasePath = ConfigurationManager.AppSettings["BReportTemplateBasePath"]
                ?? @"C:\app_data\b-templates";
        }

        [HttpPost]
        [Route("pdf")]
        public HttpResponseMessage GeneratePdf([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = GetCorrelationId();
            Log.Info($"帳票作成開始. correlationId={correlationId}");

            if (request == null)
            {
                Log.Warn($"帳票作成失敗（ボディが空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");
            }

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) ||
                !IsSafeFileNameWithExtension(request.TemplateFileName.Trim(), ".xlsx"))
            {
                Log.Warn($"帳票作成失敗（templateFileNameが不正）. correlationId={correlationId}, templateFileName='{request.TemplateFileName}'");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。");
            }

            var templatePath = Path.Combine(_templateBasePath, request.TemplateFileName);
            if (!File.Exists(templatePath))
            {
                Log.Warn($"帳票作成失敗（テンプレート未存在）. correlationId={correlationId}, templatePath='{templatePath}'");
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "テンプレートファイルが見つかりません。");
            }

            var merged = MergeToGemBoxData(request);
            var picturesMap = BuildPicturesDictionary(request);
            if (merged.Count == 0 && picturesMap.Count == 0)
            {
                Log.Warn($"帳票作成失敗（不正なリクエスト: data/tables/pictures が空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "印刷データが指定されていません。data / tables / pictures のいずれかに値を指定してください。");
            }

            Stream pdfStream;
            try
            {
                pdfStream = _pdfService.GeneratePdf(templatePath, merged, picturesMap);
            }
            catch (Exception ex)
            {
                Log.Error($"帳票作成失敗（例外）. correlationId={correlationId}", ex);
                throw;
            }

            // PDF をストリームで返却（バイト配列に全読み込みしない）。
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(pdfStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            // ファイル名はクライアント（フロント）側で決める運用のため、Content-Disposition / filename は付けない。
            Log.Info($"帳票作成完了. correlationId={correlationId}, template='{request.TemplateFileName}'");
            return response;
        }

        private string GetCorrelationId()
        {
            try
            {
                if (Request != null && Request.Headers != null &&
                    Request.Headers.TryGetValues("X-Correlation-Id", out var values))
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
        private static Dictionary<string, object> MergeToGemBoxData(GemBoxPrintRequestDto request)
        {
            var merged = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            if (request.Data != null)
            {
                foreach (var kv in request.Data)
                    merged[kv.Key] = NormalizeValue(kv.Value);
            }

            if (request.Tables != null)
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
        private static Dictionary<string, string> BuildPicturesDictionary(GemBoxPrintRequestDto request)
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

        private static object NormalizeValue(object v)
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
        private static bool IsSafeFileNameWithExtension(string name, string extensionWithDot)
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
