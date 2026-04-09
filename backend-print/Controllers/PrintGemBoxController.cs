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
            if (request == null)
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) || !IsSafeTemplateFileName(request.TemplateFileName))
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。");

            var templatePath = Path.Combine(_templateBasePath, request.TemplateFileName);
            if (!File.Exists(templatePath))
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "テンプレートファイルが見つかりません。");

            var merged = MergeToGemBoxData(request);
            if (merged == null || merged.Count == 0)
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "data / tables が空です。");

            var work = Task.Run(() => _pdfService.GeneratePdf(templatePath, merged));
            var finished = await Task.WhenAny(work, Task.Delay(TimeSpan.FromSeconds(_timeoutSeconds)));
            if (finished != work)
                return Request.CreateErrorResponse((HttpStatusCode)504, $"PDF生成がタイムアウトしました（{_timeoutSeconds}秒）。");

            var pdfStream = await work;

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(pdfStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            var fileName = string.IsNullOrWhiteSpace(request.DownloadFileName)
                ? "document.pdf"
                : Path.GetFileName(request.DownloadFileName);

            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };

            return response;
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
        /// パストラバーサル防止: ファイル名のみ、拡張子 .xlsx
        /// </summary>
        private static bool IsSafeTemplateFileName(string name)
        {
            var f = Path.GetFileName(name);
            if (!string.Equals(f, name, StringComparison.OrdinalIgnoreCase)) return false;
            if (f.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;
            return f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
        }
    }
}
