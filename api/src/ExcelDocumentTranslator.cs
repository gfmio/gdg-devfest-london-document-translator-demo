using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentTranslatorApi
{
    class ExcelDocumentTranslator : IDocumentTranslator
    {
        /// <summary>
        /// Translates an Excel document
        ///
        /// Based on method `ProcessExcelDocument` (line 424 onwards) in
        /// TranslationAssistant.Business/DocumentTranslationManager.cs in
        /// MicrosoftTranslator/DocumentTranslator
        /// </summary>
        public async Task TranslateDocument(MemoryStream memoryStream, ITextTranslator textTranslator, string to, string from = null)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, true))
            {
                List<DocumentFormat.OpenXml.Spreadsheet.Text> texts = new List<DocumentFormat.OpenXml.Spreadsheet.Text>();
                foreach (SharedStringItem si in document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
                {
                    if (si != null && si.Text != null && !String.IsNullOrEmpty(si.Text.Text))
                    {
                        texts.Add(si.Text);
                    }
                    else if (si != null)
                    {
                        texts
                            .AddRange(si.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>()
                            .Where(item => (item != null && item.Text != null && !String.IsNullOrEmpty(item.Text.Text)))
                            .Select(item => item.Text));
                    }
                }

                var textValues = texts.Select(item => item.Text);
                var translations = await textTranslator.TranslateTexts(textValues, to, from);

                using (var textsEnumerator = texts.GetEnumerator())
                {
                    using (var translationEnumerator = translations.GetEnumerator())
                    {
                        while (textsEnumerator.MoveNext() && translationEnumerator.MoveNext())
                        {
                            textsEnumerator.Current.Text = translationEnumerator.Current;
                        }
                    }
                }

                // Refresh all the shared string references.
                var tables = document.WorkbookPart.GetPartsOfType<WorksheetPart>()
                    .Select(part => part.TableDefinitionParts)
                    .SelectMany(_tables => _tables);
                foreach (var table in tables)
                {
                    foreach (TableColumn col in table.Table.TableColumns)
                    {
                        col.Name = translations.Take(int.Parse(col.Id)).Last();
                    }

                    table.Table.Save();
                }

                // Update comments
                WorkbookPart workBookPart = document.WorkbookPart;
                List<DocumentFormat.OpenXml.Spreadsheet.Comment> comments = new List<DocumentFormat.OpenXml.Spreadsheet.Comment>();
                foreach (var commentsPart in workBookPart.WorksheetParts.SelectMany(sheet => sheet.GetPartsOfType<WorksheetCommentsPart>()))
                {
                    comments.AddRange(commentsPart.Comments.CommentList.Cast<Comment>());
                }

                var commentValues = comments.Select(item => item.InnerText).ToArray();
                var translatedComments = await textTranslator.TranslateTexts(commentValues, to, from);

                using (var commentsEnumerator = comments.GetEnumerator())
                {
                    using (var translationEnumerator = translations.GetEnumerator())
                    {
                        while (commentsEnumerator.MoveNext() && translationEnumerator.MoveNext())
                        {
                            commentsEnumerator.Current.CommentText = new CommentText
                            {
                                Text = new DocumentFormat.OpenXml.Spreadsheet.Text
                                {
                                    Text = translationEnumerator.Current
                                }
                            };
                        }
                    }
                }
            }
        }
    }
}
