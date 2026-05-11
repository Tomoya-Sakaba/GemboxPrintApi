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
    /// GemBox: テンプレExcelへデータを埋め込んだ <c>.xlsx</c> を返す（PDF 化しない）。
    /// POST /api/print/gembox/excel
    /// </summary>
    [RoutePrefix("api/print/gembox")]
    public class PrintGemBoxExcelController : ApiController
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(PrintGemBoxExcelController));

        private readonly GemBoxPdfGenerationService _pdfService;
        private readonly string _templateBasePath;

        public PrintGemBoxExcelController()
        {
            _pdfService = new GemBoxPdfGenerationService();
            var configured = DbKeyValueConfig.GetRequiredString("BReportTemplateBasePath");
            _templateBasePath = PathResolveUtils.ResolveTemplateBasePath(configured);
        }

        [HttpPost]
        [Route("excel")]
        public HttpResponseMessage GenerateExcel([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = PrintGemBoxRequestUtils.GetCorrelationId(Request);
            Log.Info($"帳票Excel（埋め込み）開始. correlationId={correlationId}");

            if (request == null)
            {
                Log.Warn($"帳票Excel失敗（ボディが空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");
            }

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) ||
                !PrintGemBoxRequestUtils.IsSafeFileNameWithExtension(request.TemplateFileName.Trim(), ".xlsx"))
            {
                Log.Warn(
                    $"帳票Excel失敗（templateFileNameが不正）. correlationId={correlationId}, templateFileName='{request.TemplateFileName}'");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。");
            }

            var templatePath = Path.Combine(_templateBasePath, request.TemplateFileName.Trim());
            if (!File.Exists(templatePath))
            {
                Log.Warn($"帳票Excel失敗（テンプレート未存在）. correlationId={correlationId}, templatePath='{templatePath}'");
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "テンプレートファイルが見つかりません。");
            }

            var merged = PrintGemBoxRequestUtils.MergeToGemBoxData(request);
            var picturesMap = PrintGemBoxRequestUtils.BuildPicturesDictionary(request);
            if (merged.Count == 0 && picturesMap.Count == 0)
            {
                Log.Warn($"帳票Excel失敗（data/tables/pictures が空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "印刷データが指定されていません。data / tables / pictures のいずれかに値を指定してください。");
            }

            Stream excelStream;
            try
            {
                excelStream = _pdfService.GenerateFilledExcel(templatePath, merged, picturesMap);
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
            Log.Info($"帳票Excel（埋め込み）完了. correlationId={correlationId}, template='{request.TemplateFileName}'");
            return response;
        }
    }
}
