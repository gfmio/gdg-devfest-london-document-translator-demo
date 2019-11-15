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
                    // Find all text items, notes and comments in all slides
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

                    // Translate and replace the text items
                    await ReplaceTextsWithTranslation(texts, textTranslator, to, from);
                    // Translate and replace the notes
                    await ReplaceTextsWithTranslation(notes, textTranslator, to, from);

                    // Translate and replace the comments
                    if (comments.Count() > 0)
                    {
                        // Extract text from comment for translation
                        var values = comments.Select(text => text.InnerText);

                        // Do translation
                        var translatedComments = await textTranslator.TranslateTexts(values, to, from);

                        // Apply translations to document by iterating through both lists and
                        // replacing the original comment text with its translation
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
                // Extract text for translation
                var values = texts.Select(text => text.Text);

                // Do translation
                var translations = await textTranslator.TranslateTexts(values, to, from);

                // Apply translations to document by iterating through both lists and
                // replacing the original text with its translation
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
        /// Extracts the text content from a slide
        ///
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
