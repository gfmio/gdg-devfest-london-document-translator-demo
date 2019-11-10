using System.Threading.Tasks;

namespace DocumentTranslatorApi
{
    interface ITextTranslator
    {
        int MaxRequestSize { get; }
        int MaxElements { get; }
        Task<string[]> TranslateTextArray(string[] texts, string to, string from = null);
        Task<ILanguage[]> ListSupportedLanguages();
    }
}
