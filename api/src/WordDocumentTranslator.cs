using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentTranslatorApi
{
    class WordDocumentTranslator : IDocumentTranslator
    {
        Task IDocumentTranslator.TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from)
        {
            return TranslateDocument(memoryStream, textTranslator, to, from);
        }

        /// <summary>
        /// Translates an Word document
        ///
        /// Based on method `ProcessWordDocument` (line 726 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        public async Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null, bool ignoreHidden = false)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                // Simply the Word document mark-up
                OpenXmlPowerTools.SimplifyMarkupSettings settings = new OpenXmlPowerTools.SimplifyMarkupSettings
                {
                    AcceptRevisions = true,
                    NormalizeXml = true, //setting this to false reduces translation quality, but if true some documents have XML format errors when opening
                    RemoveBookmarks = true,
                    RemoveComments = true,
                    RemoveContentControls = true,
                    RemoveEndAndFootNotes = true,
                    RemoveFieldCodes = true,
                    RemoveGoBackBookmark = true,
                    //RemoveHyperlinks = false,
                    RemoveLastRenderedPageBreak = true,
                    RemoveMarkupForDocumentComparison = true,
                    RemovePermissions = false,
                    RemoveProof = true,
                    RemoveRsidInfo = true,
                    RemoveSmartTags = true,
                    RemoveSoftHyphens = true,
                    RemoveWebHidden = true,
                    ReplaceTabsWithSpaces = false
                };
                OpenXmlPowerTools.MarkupSimplifier.SimplifyMarkup(doc, settings);
            }

            var texts = new List<Text>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                // Find all text nodes in the document (body, headers & footers)
                var body = doc.MainDocumentPart.Document.Body;
                texts.AddRange(body.Descendants<Text>()
                    .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

                var headers = doc.MainDocumentPart.HeaderParts.Select(p => p.Header);
                foreach (var header in headers)
                {
                    texts.AddRange(header.Descendants<Text>().Where(text =>
                        !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));
                }

                var footers = doc.MainDocumentPart.FooterParts.Select(p => p.Footer);
                foreach (var footer in footers)
                {
                    texts.AddRange(footer.Descendants<Text>().Where(text =>
                        !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));
                }

                if (ignoreHidden)
                {
                    texts.RemoveAll(t => t.Parent.Descendants<Vanish>().Any());
                }

                // Extract text strings for translation
                var values = texts.Select(text => text.Text).ToArray();

                // Do Translation
                var translations = await textTranslator.TranslateTexts(values, to, from);

                // Apply translations to document by iterating through both lists and
                // replacing the original with its translation
                using (var textsEnumerator = texts.GetEnumerator())
                {
                    using (var translationsEnumerator = translations.GetEnumerator())
                    {
                        while (textsEnumerator.MoveNext() && translationsEnumerator.MoveNext())
                        {
                            textsEnumerator.Current.Text = translationsEnumerator.Current;
                        }
                    }
                }
            }
        }
    }
}
