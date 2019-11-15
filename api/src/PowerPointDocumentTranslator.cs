using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocumentTranslatorApi
{
    class PowerPointDocumentTranslator : IDocumentTranslator
    {
        /// <summary>
        /// Translates an PowerPoint document
        ///
        /// Based on method `ProcessPowerPointDocument` (line 569 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        public async Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null)
        {
            using (PresentationDocument doc = PresentationDocument.Open(memoryStream, true))
            {
                var texts = new List<DocumentFormat.OpenXml.Drawing.Text>();
                var notes = new List<DocumentFormat.OpenXml.Drawing.Text>();
                var comments = new List<Comment>();

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
                                comments.AddRange(commentsPart.CommentList.Cast<Comment>());
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

                    if (comments.Count() > 0)
                    {
                        // Extract Text for Translation
                        var values = comments.Select(text => text.InnerText);

                        // Do translation
                        var translatedComments = await textTranslator.TranslateTexts(values, to, from);

                        // Apply translations to document
                        using (var commentsEnumerator = comments.GetEnumerator())
                        {
                            using (var translationsEnumerator = translatedComments.GetEnumerator())
                            {
                                while (commentsEnumerator.MoveNext() && translationsEnumerator.MoveNext())
                                {
                                    commentsEnumerator.Current.Text = new DocumentFormat.OpenXml.Presentation.Text
                                    {
                                        Text = translationsEnumerator.Current
                                    };
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Based on method `ReplaceTextsWithTranslation` (line 667 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        private static async Task ReplaceTextsWithTranslation(List<DocumentFormat.OpenXml.Drawing.Text> texts, ITextTranslator textTranslator, string to, string from)
        {
            if (texts.Count() > 0)
            {
                // Extract Text for Translation
                var values = texts.Select(text => text.Text);

                // Do translation
                var translations = await textTranslator.TranslateTexts(values, to, from);

                // Apply translated batch to document
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

        /// <summary>
        /// Based on method `ExtractTextContent` (line 718 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        private static void ExtractTextContent(List<DocumentFormat.OpenXml.Drawing.Text> textList, DocumentFormat.OpenXml.OpenXmlElement element)
        {
            foreach (var para in element.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                textList.AddRange(
                    para.Elements<DocumentFormat.OpenXml.Drawing.Run>()
                        .Where(item => (
                            item != null &&
                            item.Text != null &&
                            !String.IsNullOrEmpty(item.Text.Text)))
                        .Select(item => item.Text));
            }
        }
    }
}
