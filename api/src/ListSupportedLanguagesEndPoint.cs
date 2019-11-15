using System;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace DocumentTranslatorApi
{
    public static class ListSupportedLanguagesEndPoint
    {
        [FunctionName("supported-languages")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            try
            {
                var translator = new GoogleTextTranslator();
                var languages = await translator.ListSupportedLanguages();
                return new OkObjectResult(languages);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return new ObjectResult(new { error }) { StatusCode = 500 };
            }
        }
    }
}
