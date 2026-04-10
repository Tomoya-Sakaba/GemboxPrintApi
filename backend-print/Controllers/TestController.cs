using System;
using System.Web.Http;

namespace backend_print.Controllers
{
    /// <summary>
    /// 疎通確認用の追加エンドポイント（hello以外の例）。
    ///
    /// GET /api/test
    /// </summary>
    [RoutePrefix("api/test")]
    public class TestController : ApiController
    {
        [HttpGet]
        [Route("")]
        public IHttpActionResult Get()
        {
            return Ok(new
            {
                message = "Test from backend-print!",
                serverTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });
        }
    }
}

