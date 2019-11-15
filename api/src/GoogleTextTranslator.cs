using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text;

using Google.Cloud.Translation.V2;

namespace DocumentTranslatorApi
{
    class GoogleTextTranslator : ITextTranslator
    {
        private readonly TranslationClient client;

        private static readonly int MAX_REQUEST_SIZE = 5000;
        private static readonly int MAX_ELEMENTS = 25;

        internal GoogleTextTranslator()
        {
            this.client = TranslationClient.Create();
        }

        public async Task<IEnumerable<string>> TranslateTexts(IEnumerable<string> texts, string to, string from = null)
        {
            var batches = SplitList(texts, MAX_ELEMENTS, MAX_REQUEST_SIZE);
            var exceptions = new Queue<Exception>();
            var translations = new List<string>();

            // Iterate through the batches, one-by-one to avoid rate limiting
            for (var l = 0; l < batches.Count(); l++)
            {
                try
                {
                    var translatedBatch = await TranslateTextBatch(
                        batches[l],
                        to,
                        from);
                    translations.AddRange(translatedBatch);
                } catch (Exception e)
                {
                    exceptions.Enqueue(e);
                }
            }

            // Throw the exceptions here after the loop completes.
            if (exceptions.Count > 0)
            {
                throw new AggregateException(exceptions);
            }

            return translations;
        }

        public async Task<IEnumerable<ILanguage>> ListSupportedLanguages()
        {
            var languages = await client.ListLanguagesAsync(target: "en");
            return languages
                .Select((language) =>
                    new Language
                    {
                        Name = language.Name,
                        Code = language.Code
                    }
                );
        }

        private async Task<IEnumerable<string>> TranslateTextBatch(IEnumerable<string> texts, string to, string from = null)
        {
            // We filter out empty strings and null values to prevent errors from Google Cloud Translation
            var nonEmptyTexts = texts.Where((text) => text != null && text.Length > 0);

            IEnumerable<string> translatedTexts;

            // If the languages are the same, we don't need to translate
            if (to == from)
            {
                translatedTexts = nonEmptyTexts;
            }
            // If there are no non-empty texts, we don't need to translate
            else if (nonEmptyTexts.Count() == 0)
            {
                translatedTexts = new string[0];
            }
            else
            {
                // If from is null, Google Cloud Translation will identify the language automatically
                var translateTextResult = await client.TranslateTextAsync(nonEmptyTexts, to, from);
                // Extract the strings
                translatedTexts = translateTextResult.Select((item) => item.TranslatedText);
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
                        var translation = translatedTexts.Take(i + 1).Last();
                        i++;
                        return translation;
                    }
                });
        }

        /// <summary>
        /// Splits the list.
        ///
        /// Based on method `SplitList` (line 851 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        /// </summary>
        /// <param name="values">
        ///  The values to be split.
        /// </param>
        /// <param name="groupSize">
        ///  The group size.
        /// </param>
        /// <param name="maxSize">
        ///  The max size.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        ///  The System.Collections.Generic.List`1[T -&gt; System.Collections.Generic.List`1[T -&gt; T]].
        /// </returns>
        private static List<List<T>> SplitList<T>(IEnumerable<T> values, int groupSize, int maxSize)
        {
            List<List<T>> result = new List<List<T>>();
            List<T> valueList = values.ToList();
            int startIndex = 0;
            int count = valueList.Count;

            while (startIndex < count)
            {
                int elementCount = (startIndex + groupSize > count) ? count - startIndex : groupSize;
                while (true)
                {
                    var aggregatedSize =
                        valueList.GetRange(startIndex, elementCount)
                            .Aggregate(
                                new StringBuilder(),
                                (s, i) => s.Length < maxSize ? s.Append(i) : s,
                                s => s.ToString())
                            .Length;
                    if (aggregatedSize >= maxSize)
                    {
                        if (elementCount == 1) break;
                        elementCount = elementCount - 1;
                    }
                    else
                    {
                        break;
                    }
                }

                result.Add(valueList.GetRange(startIndex, elementCount));
                startIndex += elementCount;
            }

            return result;
        }
    }
}
