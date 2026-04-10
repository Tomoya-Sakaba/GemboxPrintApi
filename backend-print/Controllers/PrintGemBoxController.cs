using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using backend_print.Models.DTOs;
using backend_print.Services;
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
        private readonly GemBoxPdfGenerationService _pdfService;
        private readonly string _templateBasePath;
        private readonly int _timeoutSeconds;

        public PrintGemBoxController()
        {
            _pdfService = new GemBoxPdfGenerationService();
            _templateBasePath = ConfigurationManager.AppSettings["BReportTemplateBasePath"]
                ?? @"C:\app_data\b-templates";
            _timeoutSeconds = int.TryParse(ConfigurationManager.AppSettings["GemBoxPdfTimeoutSeconds"], out var s)
                ? s
                : 60;
        }

        [HttpPost]
        [Route("pdf")]
        public async Task<HttpResponseMessage> GeneratePdf([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = GetCorrelationId();
            SimpleFileLogger.Log(
                ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                $"[api] GeneratePdf start. correlationId={correlationId}");

            if (request == null)
            {
                SimpleFileLogger.Log(
                    ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                    $"[api] GeneratePdf bad request (null body). correlationId={correlationId}");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");
            }

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) ||
                !IsSafeFileNameWithExtension(request.TemplateFileName.Trim(), ".xlsx"))
            {
                SimpleFileLogger.Log(
                    ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                    $"[api] GeneratePdf bad request (templateFileName). correlationId={correlationId}, templateFileName='{request.TemplateFileName}'");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。");
            }

            var templatePath = Path.Combine(_templateBasePath, request.TemplateFileName);
            if (!File.Exists(templatePath))
            {
                SimpleFileLogger.Log(
                    ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                    $"[api] GeneratePdf not found (template). correlationId={correlationId}, templatePath='{templatePath}'");
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "テンプレートファイルが見つかりません。");
            }

            var merged = MergeToGemBoxData(request);
            if (merged == null || merged.Count == 0)
            {
                SimpleFileLogger.Log(
                    ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                    $"[api] GeneratePdf bad request (empty data/tables). correlationId={correlationId}");
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "印刷データが指定されていません。data もしくは tables のどちらかに値を指定してください。");
            }

            // GeneratePdf は CPU/IO が重くなり得るため、タイムアウト付きで別タスクとして実行する。
            var work = Task.Run(() => _pdfService.GeneratePdf(templatePath, merged));
            var finished = await Task.WhenAny(work, Task.Delay(TimeSpan.FromSeconds(_timeoutSeconds)));
            if (finished != work)
            {
                SimpleFileLogger.Log(
                    ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                    $"[api] GeneratePdf timeout. correlationId={correlationId}, timeoutSeconds={_timeoutSeconds}");
                return Request.CreateErrorResponse((HttpStatusCode)504, $"PDF生成がタイムアウトしました（{_timeoutSeconds}秒）。");
            }

            var pdfStream = await work;

            // PDF をストリームで返却（バイト配列に全読み込みしない）。
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(pdfStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            // ダウンロード名: ファイル名のみ採用（パスは落とす）
            var fileName = Path.GetFileName((request.DownloadFileName ?? "").Trim().Trim('"'));
            if (string.IsNullOrWhiteSpace(fileName) || !IsSafeFileNameWithExtension(fileName, ".pdf"))
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "downloadFileName が未指定または不正です。ファイル名のみ（例: example.pdf）を指定してください。");

            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };

            SimpleFileLogger.Log(
                ConfigurationManager.AppSettings["GemBoxLogFilePath"],
                $"[api] GeneratePdf ok. correlationId={correlationId}, fileName='{fileName}', template='{request.TemplateFileName}'");
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
        /// data（単票）と tables（明細）を GemBoxPdfGenerationService が期待する1つの Dictionary にまとめる。
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
