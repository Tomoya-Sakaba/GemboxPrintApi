using System;
using System.Web.Http;

namespace backend_print.Controllers
{
    /// <summary>
    /// 疎通確認用の最小エンドポイント。
    /// - ExcelテンプレやPDF生成などの依存を一切通さず
    /// - 「backend-print が起動していて HTTP で到達できる」ことだけを確認する目的
    ///
    /// GET /api/hello
    /// </summary>
    [RoutePrefix("api/hello")]
    public class HelloController : ApiController
    {
        [HttpGet]
        [Route("")]
        public IHttpActionResult Get()
        {
            // 返却は匿名型でOK（Web API が JSON にシリアライズする）。
            // serverTime を返すことで「レスポンスがキャッシュではなく backend-print の実行結果」だと分かりやすい。
            return Ok(new
            {
                message = "Hello from backend-print!",
                serverTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });
        }
    }
}

