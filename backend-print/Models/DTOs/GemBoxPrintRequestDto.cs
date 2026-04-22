using System.Collections.Generic;
using Newtonsoft.Json;

namespace backend_print.Models.DTOs
{
    /// <summary>
    /// 汎用GemBox印刷リクエスト（backend から POST で受け取る）
    /// </summary>
    public class GemBoxPrintRequestDto
    {
        [JsonProperty("templateFileName")]
        public string TemplateFileName { get; set; }

        [JsonProperty("data")]
        public Dictionary<string, object> Data { get; set; }

        [JsonProperty("tables")]
        public Dictionary<string, List<Dictionary<string, object>>> Tables { get; set; }

        [JsonProperty("pictures")]
        public Dictionary<string, string> Pictures { get; set; }
    }
}
