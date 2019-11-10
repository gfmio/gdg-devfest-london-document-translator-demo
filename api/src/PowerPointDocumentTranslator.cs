using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentTranslatorApi
{
    class PowerPointDocumentTranslator : IDocumentTranslator
    {
        /// <summary>
        /// Based on method `ProcessPowerPointDocument` in TranslationAssistant.Business/DocumentTranslationManager.cs line 569 onwards
        /// </summary>
        public async Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null)
        {
            using (PresentationDocument doc = PresentationDocument.Open(memoryStream, true))
            {
                List<DocumentFormat.OpenXml.Drawing.Text> texts = new List<DocumentFormat.OpenXml.Drawing.Text>();
                List<DocumentFormat.OpenXml.Drawing.Text> notes = new List<DocumentFormat.OpenXml.Drawing.Text>();
                List<DocumentFormat.OpenXml.Presentation.Comment> lstComments = new List<DocumentFormat.OpenXml.Presentation.Comment>();

                var slideParts = doc.PresentationPart.SlideParts;
                if (slideParts != null)
                {
                    foreach (var slidePart in slideParts)
                    {
                        if (slidePart.Slide != null)
                        {
                            var slide = slidePart.Slide;
                            ExtractTextContent(texts, slide);

                            var commentsPart = slidePart.SlideCommentsPart;
                            if (commentsPart != null)
                            {
                                lstComments.AddRange(commentsPart.CommentList.Cast<DocumentFormat.OpenXml.Presentation.Comment>());
                            }

                            var notesPart = slidePart.NotesSlidePart;
                            if (notesPart != null)
                            {
                                ExtractTextContent(notes, notesPart.NotesSlide);
                            }
                        }
                    }

                    await ReplaceTextsWithTranslation(texts, textTranslator, to, from);
                    await ReplaceTextsWithTranslation(notes, textTranslator, to, from);

                    if (lstComments.Count() > 0)
                    {
                        // Extract Text for Translation
                        var batch = lstComments.Select(text => text.InnerText);

                        // Do Translation
                        var batchesComments = Splitter.SplitList(batch, textTranslator.MaxElements, textTranslator.MaxRequestSize);

                        // Use ConcurrentQueue to enable safe enqueueing from multiple threads. 
                        var exceptions = new ConcurrentQueue<Exception>();

                        for (var l = 0; l < batchesComments.Count(); l++)
                        {
                            try
                            {
                                var translationOutput =
                                    await textTranslator.TranslateTextArray(
                                        batchesComments[l].ToArray(),
                                        to,
                                        from);
                                int batchStartIndexInDocument = 0;
                                for (int i = 0; i < l; i++)
                                {
                                    batchStartIndexInDocument = batchStartIndexInDocument
                                                                + batchesComments[i].Count();
                                }

                                // Apply translated batch to document
                                for (int j = 0; j < translationOutput.Length; j++)
                                {
                                    int indexInDocument = j + batchStartIndexInDocument + 1;
                                    var newValue = translationOutput[j];
                                    var commentPart = lstComments.Take(indexInDocument).Last();
                                    commentPart.Text = new DocumentFormat.OpenXml.Presentation.Text
                                    {
                                        Text = newValue
                                    };
                                }
                            }
                            catch (Exception ex)
                            {
                                exceptions.Enqueue(ex);
                            }
                        }

                        // Throw the exceptions here after the loop completes. 
                        if (exceptions.Count > 0)
                        {
                            throw new AggregateException(exceptions);
                        }
                    }
                }
            }

        }

        /// Based on method `ReplaceTextsWithTranslation` in TranslationAssistant.Business/DocumentTranslationManager.cs line 667 onwards
        private static async Task ReplaceTextsWithTranslation(List<DocumentFormat.OpenXml.Drawing.Text> texts, ITextTranslator textTranslator, string to, string from)
        {
            if (texts.Count() > 0)
            {
                // Extract Text for Translation
                var batch = texts.Select(text => text.Text);

                // Do Translation
                var batches = Splitter.SplitList(batch, textTranslator.MaxElements, textTranslator.MaxRequestSize);

                var exceptions = new Queue<Exception>();

                for (var l = 0; l < batches.Count(); l++)
                {
                    try
                    {
                        var translationOutput = await textTranslator.TranslateTextArray(batches[l].ToArray(), to, from);
                        int batchStartIndexInDocument = 0;
                        for (int i = 0; i < l; i++)
                        {
                            batchStartIndexInDocument = batchStartIndexInDocument
                                                        + batches[i].Count();
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
                }

                // Throw the exceptions here after the loop completes. 
                if (exceptions.Count > 0)
                {
                    throw new AggregateException(exceptions);
                }
            }
        }

        /// Based on method `ExtractTextContent` in TranslationAssistant.Business/DocumentTranslationManager.cs line 718 onwards
        private static void ExtractTextContent(List<DocumentFormat.OpenXml.Drawing.Text> textList, DocumentFormat.OpenXml.OpenXmlElement element)
        {
            foreach (DocumentFormat.OpenXml.Drawing.Paragraph para in element.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                textList.AddRange(para.Elements<DocumentFormat.OpenXml.Drawing.Run>().Where(item => (item != null && item.Text != null && !String.IsNullOrEmpty(item.Text.Text))).Select(item => item.Text));
            }
        }
    }
}
