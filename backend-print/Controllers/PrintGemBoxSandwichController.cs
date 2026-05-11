using System;
using System.Configuration;
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
    /// GemBox: （パターンB）1つのExcelブックの「前半シート→中間PDF→後半シート」を1本に結合。
    /// POST /api/print/gembox/sandwich-pdf
    /// </summary>
    [RoutePrefix("api/print/gembox")]
    public class PrintGemBoxSandwichController : ApiController
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(PrintGemBoxSandwichController));

        private readonly GemBoxPdfGenerationService _pdfService;
        private readonly string _templateBasePath;

        public PrintGemBoxSandwichController()
        {
            _pdfService = new GemBoxPdfGenerationService();
            var configured = DbKeyValueConfig.GetRequiredString("BReportTemplateBasePath");
            _templateBasePath = PathResolveUtils.ResolveTemplateBasePath(configured);
        }

        [HttpPost]
        [Route("sandwich-pdf")]
        public HttpResponseMessage GenerateSandwichPdf([FromBody] GemBoxPrintSandwichPdfRequestDto request)
        {
            var correlationId = PrintGemBoxRequestUtils.GetCorrelationId(Request);
            Log.Info($"サンドイッチPDF開始. correlationId={correlationId}");

            if (request == null)
            {
                Log.Warn($"サンドイッチPDF失敗（ボディが空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "リクエストボディが空です。");
            }

            if (!TryResolveMiddlePdfPath(request, correlationId, out var middlePath, out var middleErr))
                return middleErr;

            if (string.IsNullOrWhiteSpace(request.TemplateFileName) ||
                !PrintGemBoxRequestUtils.IsSafeFileNameWithExtension(request.TemplateFileName.Trim(), ".xlsx"))
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "templateFileName が不正です（ファイル名のみ、.xlsx を指定）。");
            }

            if (!request.FirstSheetIndex.HasValue || !request.SecondSheetIndex.HasValue)
            {
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "firstSheetIndex / secondSheetIndex を指定してください（0始まり）。");
            }

            var templatePath = Path.Combine(_templateBasePath, request.TemplateFileName.Trim());
            if (!File.Exists(templatePath))
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "テンプレートファイルが見つかりません。");

            var i0 = request.FirstSheetIndex.Value;
            var i1 = request.SecondSheetIndex.Value;
            if (i0 < 0 || i1 < 0)
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "シートindexは0以上を指定してください。");

            return RunMergeOneBook(correlationId, templatePath, i0, i1, middlePath, request);
        }

        /// <summary>
        /// 中間PDFの物理パス（フルパスのみ）を検証する。<paramref name="middlePath"/> は解決できたときだけ設定。
        /// </summary>
        private bool TryResolveMiddlePdfPath(
            GemBoxPrintSandwichPdfRequestDto request,
            string correlationId,
            out string middlePath,
            out HttpResponseMessage errorResponse)
        {
            middlePath = null;
            errorResponse = null;

            var pathRaw = (request.MiddlePdfPath ?? "").Trim();
            if (string.IsNullOrEmpty(pathRaw))
            {
                Log.Warn($"サンドイッチPDF: 中間PDF解決失敗（middlePdfPath 未指定）. correlationId={correlationId}");
                errorResponse = Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "middlePdfPath に中間PDFのフルパスを指定してください。");
                return false;
            }

            if (!Path.IsPathRooted(pathRaw))
            {
                Log.Warn($"サンドイッチPDF: 中間PDF解決失敗（フルパスではない）. correlationId={correlationId}, middlePdfPath='{pathRaw}'");
                errorResponse = Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "middlePdfPath はローカル絶対パス（フルパス）で指定してください。");
                return false;
            }

            var allow = (ConfigurationManager.AppSettings["GemBoxSandwichAllowAbsoluteMiddlePdf"] ?? "").Trim();
            if (!string.Equals(allow, "true", StringComparison.OrdinalIgnoreCase))
            {
                Log.Warn(
                    $"サンドイッチPDF: 中間PDF解決失敗（絶対パスは未許可）. correlationId={correlationId}, middlePdfPath='{pathRaw}'");
                errorResponse = Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "middlePdfPath が絶対パスのときは Web.config の GemBoxSandwichAllowAbsoluteMiddlePdf=true が必要です。");
                return false;
            }

            if (!File.Exists(pathRaw))
            {
                Log.Warn(
                    $"サンドイッチPDF: 中間PDF解決失敗（ファイルなし）. correlationId={correlationId}, path='{pathRaw}'");
                errorResponse = Request.CreateErrorResponse(HttpStatusCode.NotFound, "中間PDFファイルが見つかりません。");
                return false;
            }

            middlePath = pathRaw;
            return true;
        }

        private HttpResponseMessage RunMergeOneBook(
            string correlationId,
            string templatePath,
            int firstSheetIndex,
            int secondSheetIndex,
            string middlePdfPath,
            GemBoxPrintSandwichPdfRequestDto request)
        {
            var baseReq = new GemBoxPrintRequestDto
            {
                Data = request.Data,
                Tables = request.Tables,
                Pictures = request.Pictures
            };
            var merged = PrintGemBoxRequestUtils.MergeToGemBoxData(baseReq);
            var picturesMap = PrintGemBoxRequestUtils.BuildPicturesDictionary(baseReq);
            if (merged.Count == 0 && picturesMap.Count == 0)
            {
                Log.Warn($"サンドイッチPDF失敗（data空）. correlationId={correlationId}");
                return Request.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "印刷データが指定されていません。data / tables / pictures のいずれかに値を指定してください。");
            }

            Stream pdfStream;
            try
            {
                pdfStream = _pdfService.GenerateSandwichPdfOneWorkbookTwoSheets(
                    templatePath,
                    firstSheetIndex,
                    secondSheetIndex,
                    middlePdfPath,
                    merged,
                    picturesMap);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Log.Warn($"サンドイッチPDF失敗（シートindex）. correlationId={correlationId}", ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
            catch (Exception ex)
            {
                Log.Error($"サンドイッチPDF失敗（例外）. correlationId={correlationId}", ex);
                throw;
            }

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(pdfStream)
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            Log.Info($"サンドイッチPDF完了（1ブック2シート）. correlationId={correlationId}");
            return response;
        }
    }
}
