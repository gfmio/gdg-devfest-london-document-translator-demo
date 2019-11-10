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
                List<DocumentFormat.OpenXml.Spreadsheet.Text> lstTexts = new List<DocumentFormat.OpenXml.Spreadsheet.Text>();
                foreach (SharedStringItem si in document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
                {
                    if (si != null && si.Text != null && !String.IsNullOrEmpty(si.Text.Text))
                    {
                        lstTexts.Add(si.Text);
                    }
                    else if (si != null)
                    {
                        lstTexts
                            .AddRange(si.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>()
                            .Where(item => (item != null && item.Text != null && !String.IsNullOrEmpty(item.Text.Text)))
                            .Select(item => item.Text));
                    }
                }

                var batch = lstTexts.Select(item => item.Text);
                IEnumerable<string> values = batch as string[] ?? batch.ToArray();

                var batches = Splitter.SplitList(values, textTranslator.MaxElements, textTranslator.MaxRequestSize);
                string[] translated = new string[values.Count()];

                var exceptions = new Queue<Exception>();

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
                            translated[indexInDocument - 1] = newValue;
                            lstTexts[indexInDocument - 1].Text = newValue;
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Enqueue(ex);
                    }
                }

                if (exceptions.Count > 0)
                {
                    throw new AggregateException(exceptions);
                }

                // Refresh all the shared string references.
                var tables = document.WorkbookPart.GetPartsOfType<WorksheetPart>()
                    .Select(part => part.TableDefinitionParts)
                    .SelectMany(_tables => _tables);
                foreach (var table in tables)
                {
                    foreach (TableColumn col in table.Table.TableColumns)
                    {
                        col.Name = translated[int.Parse(col.Id) - 1];
                    }

                    table.Table.Save();
                }

                // Update comments
                WorkbookPart workBookPart = document.WorkbookPart;
                List<DocumentFormat.OpenXml.Spreadsheet.Comment> lstComments = new List<DocumentFormat.OpenXml.Spreadsheet.Comment>();
                foreach (WorksheetCommentsPart commentsPart in workBookPart.WorksheetParts.SelectMany(sheet => sheet.GetPartsOfType<WorksheetCommentsPart>()))
                {
                    lstComments.AddRange(commentsPart.Comments.CommentList.Cast<Comment>());
                }

                var batchComments = lstComments.Select(item => item.InnerText);
                var batchesComments = Splitter.SplitList(batchComments, textTranslator.MaxElements, textTranslator.MaxRequestSize);
                string[] translatedComments = new string[batchesComments.Count()];

                for (var l = 0; l < batchesComments.Count(); l++)
                {
                    try
                    {
                        var translationOutput = await textTranslator.TranslateTextArray(
                                batchesComments[l].ToArray(),
                                to,
                                from);
                        int batchStartIndexInDocument = 0;
                        for (int i = 0; i < l; i++)
                        {
                            batchStartIndexInDocument = batchStartIndexInDocument + batches[i].Count();
                        }

                        for (int j = 0; j < translationOutput.Length; j++)
                        {
                            int indexInDocument = j + batchStartIndexInDocument + 1;
                            var currentSharedStringItem = lstComments.Take(indexInDocument).Last();
                            var newValue = translationOutput[j];
                            if (translatedComments.Count() > indexInDocument - 1)
                            {
                                translatedComments[indexInDocument - 1] = newValue;
                            }
                            currentSharedStringItem.CommentText = new CommentText
                            {
                                Text = new DocumentFormat.OpenXml.Spreadsheet.Text { Text = newValue }
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
