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
using backend_print.Utils;
using log4net;

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
            var configured = DbKeyValueConfig.GetRequiredString("BReportTemplateBasePath");
            _templateBasePath = PathResolveUtils.ResolveTemplateBasePath(configured);
        }

        [HttpPost]
        [Route("pdf")]
        public HttpResponseMessage GeneratePdf([FromBody] GemBoxPrintRequestDto request)
        {
            var correlationId = PrintGemBoxRequestUtils.GetCorrelationId(Request);
            Log.Info($"帳票作成開始. correlationId={correlationId}");

            if (request == null)
            {
                Log.Warn($"帳票作成失敗（ボディが空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");
            }

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) ||
                !PrintGemBoxRequestUtils.IsSafeFileNameWithExtension(request.TemplateFileName.Trim(), ".xlsx"))
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

            var merged = PrintGemBoxRequestUtils.MergeToGemBoxData(request);
            var picturesMap = PrintGemBoxRequestUtils.BuildPicturesDictionary(request);
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
    }
}
