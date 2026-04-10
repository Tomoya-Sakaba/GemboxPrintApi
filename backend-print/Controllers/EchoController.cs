using System.Web.Http;

namespace backend_print.Controllers
{
    /// <summary>
    /// POST 疎通・サンプル用。受け取った JSON をそのまま返す（insert の代わりの動作確認用）。
    /// 本番の insert は別コントローラで実装し、backend 側の HTTP プロキシから同様に転送する。
    /// </summary>
    [RoutePrefix("api/echo")]
    public class EchoController : ApiController
    {
        [HttpPost]
        [Route("")]
        public IHttpActionResult Post([FromBody] object body)
        {
            return Ok(new { echo = body, note = "backend-print received POST body" });
        }
    }
}
