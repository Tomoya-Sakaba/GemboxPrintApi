using System.Collections.Generic;
using Newtonsoft.Json;

namespace backend_print.Models.DTOs
{
    /// <summary>
    /// Excel→PDF（前半シート）＋既存PDF＋Excel→PDF（後半シート）を1本にまとめるリクエスト（パターンB専用）。
    /// </summary>
    public class GemBoxPrintSandwichPdfRequestDto
    {
        /// <summary>中間PDFのフルパス（ローカル絶対パス）。<c>GemBoxSandwichAllowAbsoluteMiddlePdf=true</c> が必要。</summary>
        [JsonProperty("middlePdfPath")]
        public string MiddlePdfPath { get; set; }

        [JsonProperty("templateFileName")]
        public string TemplateFileName { get; set; }

        [JsonProperty("firstSheetIndex")]
        public int? FirstSheetIndex { get; set; }

        [JsonProperty("secondSheetIndex")]
        public int? SecondSheetIndex { get; set; }

        [JsonProperty("data")]
        public Dictionary<string, object> Data { get; set; }

        [JsonProperty("tables")]
        public Dictionary<string, List<Dictionary<string, object>>> Tables { get; set; }

        [JsonProperty("pictures")]
        public Dictionary<string, string> Pictures { get; set; }
    }
}
