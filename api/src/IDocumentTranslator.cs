using System.IO;
using System.Threading.Tasks;

namespace DocumentTranslatorApi
{
    interface IDocumentTranslator
    {
        Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null);
    }
}
