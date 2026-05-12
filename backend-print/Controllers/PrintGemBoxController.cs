using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using backend_print.Models.DTOs;
using backend_print.Services;
using backend_print.Utils;
using log4net;

namespace backend_print.Controllers
{
    /// <summary>
    /// GemBox: テンプレExcelへデータを埋め込み、PDF または埋め込み済み <c>.xlsx</c> を返す（DBアクセスなし）。
    /// POST /api/print/gembox/pdf … PDF（任意で <c>addPdfPath</c> のPDFを末尾に結合）
    /// POST /api/print/gembox/excel … Excel
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
            var configured = DbKeyValueConfig.GetRequiredString("BReportTemplateBasePath");
            _templateBasePath = PathResolveUtils.ResolveTemplateBasePath(configured);
        }

        [HttpPost]
        [Route("pdf")]
        public HttpResponseMessage GeneratePdf([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = PrintGemBoxRequestUtils.GetCorrelationId(Request);
            Log.Info($"帳票PDF開始. correlationId={correlationId}");

            var prep = PrintGemBoxRequestUtils.TryPrepareGemBoxPrint(request, _templateBasePath);
            if (!prep.Ok)
            {
                Log.Warn($"帳票PDF {prep.FailureLogDetail}. correlationId={correlationId}");
                return Request.CreateErrorResponse(prep.ErrorStatus, prep.ErrorMessage);
            }

            var ctx = prep.Prepared;

            if (!PrintGemBoxRequestUtils.TryValidateAddPdfPathForAppend(
                    request.AddPdfPath,
                    out var appendPath,
                    out var addErrStatus,
                    out var addErrMessage,
                    out var addLogDetail))
            {
                Log.Warn($"帳票PDF {addLogDetail}. correlationId={correlationId}");
                return Request.CreateErrorResponse(addErrStatus, addErrMessage);
            }

            Stream pdfStream;
            try
            {
                byte[] mainBytes;
                using (var genStream = _pdfService.GeneratePdf(ctx.TemplatePath, ctx.MergedData, ctx.Pictures))
                using (var ms = new MemoryStream())
                {
                    genStream.CopyTo(ms);
                    mainBytes = ms.ToArray();
                }

                if (string.IsNullOrEmpty(appendPath))
                {
                    pdfStream = new MemoryStream(mainBytes) { Position = 0 };
                }
                else
                {
                    var appendBytes = File.ReadAllBytes(appendPath);
                    var merged = PdfMergeService.MergePdfs(new[] { mainBytes, appendBytes });
                    pdfStream = new MemoryStream(merged) { Position = 0 };
                }
            }
            catch (Exception ex)
            {
                Log.Error($"帳票PDF失敗（例外）. correlationId={correlationId}", ex);
                throw;
            }

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(pdfStream),
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            Log.Info($"帳票PDF完了. correlationId={correlationId}, template='{ctx.TemplateFileName}'");
            return response;
        }

        [HttpPost]
        [Route("excel")]
        public HttpResponseMessage GenerateExcel([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = PrintGemBoxRequestUtils.GetCorrelationId(Request);
            Log.Info($"帳票Excel（埋め込み）開始. correlationId={correlationId}");

            var prep = PrintGemBoxRequestUtils.TryPrepareGemBoxPrint(request, _templateBasePath);
            if (!prep.Ok)
            {
                Log.Warn($"帳票Excel {prep.FailureLogDetail}. correlationId={correlationId}");
                return Request.CreateErrorResponse(prep.ErrorStatus, prep.ErrorMessage);
            }

            var ctx = prep.Prepared;
            Stream excelStream;
            try
            {
                excelStream = _pdfService.GenerateFilledExcel(ctx.TemplatePath, ctx.MergedData, ctx.Pictures);
            }
            catch (Exception ex)
            {
                Log.Error($"帳票Excel失敗（例外）. correlationId={correlationId}", ex);
                throw;
            }

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(excelStream),
            };
            response.Content.Headers.ContentType =
                new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            Log.Info($"帳票Excel（埋め込み）完了. correlationId={correlationId}, template='{ctx.TemplateFileName}'");
            return response;
        }
    }
}
