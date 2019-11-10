using System.Linq;
using System.Threading.Tasks;

using Google.Cloud.Translation.V2;

namespace DocumentTranslatorApi
{
    class GoogleTextTranslator : ITextTranslator
    {
        private readonly TranslationClient client;

        public int MaxRequestSize { get; } = 5000;
        public int MaxElements { get; } = 25;

        internal GoogleTextTranslator()
        {
            this.client = TranslationClient.Create();
        }

        public async Task<string[]> TranslateTextArray(string[] texts, string to, string from = null)
        {
            // We filter out empty strings and null values to prevent errors from Google Cloud Translation
            var nonEmptyTexts = texts.Where((text) => text != null && text.Length > 0).ToArray();

            string[] translatedTexts;

            // If the languages are the same, we don't need to translate
            if (to == from)
            {
                translatedTexts = nonEmptyTexts;
            }
            // If there are no non-empty texts, we don't need to translate
            else if (nonEmptyTexts.Length == 0)
            {
                translatedTexts = new string[0];
            }
            else
            {
                // If from is null, Google Cloud Translation will identify the language automatically
                var translateTextResult = await client.TranslateTextAsync(nonEmptyTexts, to, from);
                // Extract the strings
                translatedTexts = translateTextResult.Select((item) => item.TranslatedText).ToArray();
            }

            // We recover an array of equal length as the original 
            // For each item in the original array, we return the empty string if it is null or empty or its translation
            var i = 0;
            return texts
                .Select((text) =>
                {
                    if (text == null || text.Length == 0)
                    {
                        return "";
                    }
                    else
                    {
                        var translation = translatedTexts[i];
                        i++;
                        return translation;
                    }
                })
                .ToArray();
        }

        public async Task<ILanguage[]> ListSupportedLanguages()
        {
            var languages = await client.ListLanguagesAsync(target: "en");
            return languages
                .Select((language) =>
                    new Language
                    {
                        Name = language.Name,
                        Code = language.Code
                    }
                )
                .ToArray();
        }

    }
}
