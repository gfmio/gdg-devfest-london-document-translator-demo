using System.Collections.Generic;
using System.Threading.Tasks;

namespace DocumentTranslatorApi
{
    interface ITextTranslator
    {
        Task<IEnumerable<string>> TranslateTexts(IEnumerable<string> texts, string to, string from = null);
        Task<IEnumerable<ILanguage>> ListSupportedLanguages();
    }
}
