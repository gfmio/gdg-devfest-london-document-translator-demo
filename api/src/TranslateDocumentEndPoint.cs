using System;
using System.IO;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Primitives;

namespace DocumentTranslatorApi
{
    public static class TranslateDocumentEndPoint
    {
        [FunctionName("translate-document")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var textTranslator = new GoogleTextTranslator();

            try
            {
                // From
                StringValues fromValues;
                req.Form.TryGetValue("from", out fromValues);
                var from = fromValues.ToString();

                // To
                StringValues toValues;
                req.Form.TryGetValue("to", out toValues);
                if (toValues.Count == 0)
                {
                    throw new Exception("Missing field `to`");
                }
                var to = toValues.ToString();

                // File
                if (req.Form.Files.Count == 0)
                {
                    throw new Exception("Missing file");
                }
                var file = req.Form.Files[0];
                var stream = file.OpenReadStream();

                // Since this stream is read-only, we create a copy that we mutate
                var memoryStream = new MemoryStream();
                stream.CopyTo(memoryStream);
                // Reset the position to the start
                memoryStream.Position = 0;

                // Select the appropriate translator
                var mimeType = file.ContentType;
                var documentTranslator = DocumentTranslatorByMimeType(mimeType);

                // Perform the translation
                // This modifies the memoryStream in place, so we don't get a return value back.
                await documentTranslator.TranslateDocument(memoryStream, textTranslator, to, from);

                // Return the new file
                return new FileContentResult(memoryStream.ToArray(), mimeType);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return new ObjectResult(new { error = error }) { StatusCode = 500 };
            }
        }

        private static IDocumentTranslator DocumentTranslatorByMimeType(string mimeType)
        {
            switch (mimeType)
            {
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    return new WordDocumentTranslator();
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    return new ExcelDocumentTranslator();
                case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
                    return new PowerPointDocumentTranslator();
                default:
                    throw new Exception("unsupported mime type");
            }
        }
    }
}
