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
        /// Based on method `ProcessWordDocument` in TranslationAssistant.Business/DocumentTranslationManager.cs line 726 onwards
        /// </summary>
        public async Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null, bool ignoreHidden = false)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {

                OpenXmlPowerTools.SimplifyMarkupSettings settings = new OpenXmlPowerTools.SimplifyMarkupSettings
                {
                    AcceptRevisions = true,
                    NormalizeXml = true,         //setting this to false reduces translation quality, but if true some documents have XML format errors when opening
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

            List<DocumentFormat.OpenXml.Wordprocessing.Text> texts = new List<DocumentFormat.OpenXml.Wordprocessing.Text>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                texts.AddRange(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

                var headers = doc.MainDocumentPart.HeaderParts.Select(p => p.Header);
                foreach (var header in headers)
                {
                    texts.AddRange(header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));
                }

                var footers = doc.MainDocumentPart.FooterParts.Select(p => p.Footer);
                foreach (var footer in footers)
                {
                    texts.AddRange(footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));
                }

                if (ignoreHidden)
                {
                    texts.RemoveAll(t => t.Parent.Descendants<Vanish>().Any());
                }

                var exceptions = new Queue<Exception>();

                // Extract Text for Translation
                var batch = texts.Select(text => text.Text);

                // Do Translation
                var batches = Splitter.SplitList(batch, textTranslator.MaxElements, textTranslator.MaxRequestSize);

                for (var l = 0; l < batches.Count(); l++)
                {
                    try
                    {
                        var translationOutput = await textTranslator.TranslateTextArray(
                            batches[l].ToArray(),
                            to,
                            from);
                        int batchStartIndexInDocument = 0;
                        for (int i = 0; i < l; i++)
                        {
                            batchStartIndexInDocument = batchStartIndexInDocument + batches[i].Count();
                        }

                        // Apply translated batch to document
                        for (int j = 0; j < translationOutput.Length; j++)
                        {
                            int indexInDocument = j + batchStartIndexInDocument + 1;
                            var newValue = translationOutput[j];
                            texts.Take(indexInDocument).Last().Text = newValue;
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Enqueue(ex);
                    }
                };

                // Throw the exceptions here after the loop completes. 
                if (exceptions.Count > 0)
                {
                    throw new AggregateException(exceptions);
                }
            }

        }
    }
}
