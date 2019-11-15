using System.Threading.Tasks;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace DocumentTranslatorApi
{
    public static class HelloWorldEndPoint
    {
        [FunctionName("hello-world")]
        public static Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            return Task.FromResult<IActionResult>(new OkObjectResult(new { message = "Hello World" }));
        }
    }
}
