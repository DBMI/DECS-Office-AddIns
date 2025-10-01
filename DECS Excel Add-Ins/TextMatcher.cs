using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class TextMatcher
    {
        private const int MIN_SIZE = 15;
        private Microsoft.Office.Interop.Excel.Application application;
        private Range artSparkIdColumn;
        private Range artSparkTextColumn;
        private Range redcapIdColumn;
        private Range redcapTextColumn;

        private Range crossReferenceRow;

        private Regex chunkExtractor;
        private const string chunkPattern = @"(\>[^\<\>]{3,}\<)+";
        private Regex idExtractor;
        private const string idPattern = @"(?:icr)*(?:nc)*(?:message)*(?<id>[\d\w]*)(?:message)*";
        //private Regex idExtractorAlt;
        //private const string idPatternAlt = @"nc(?<id>[\d\w]*)message";
        private Regex providerExtractor;
        private const string providerPattern = @"Provider reply:<\/strong><br><br \/><span style=""font-weight: normal;""(?<provider_reply>.*)";
        private Regex replyExtractor;
        private const string replyPattern = @"\>(?<reply>(\d|\s|\w|,|\.|\<\d|(?<!\w)\>|\/|\(|\)|\[|\]|'|""|\?|!|:|\*|\-)*)";

        private List<string> redcapMessagesMatched;

        private void BuildCrossReferenceSheet()
        {
            Worksheet crossReferenceSheet = Utilities.CreateNewNamedSheet("Cross Reference");
            System.Threading.Thread.Sleep(500);
            crossReferenceRow = crossReferenceSheet.Range["A1"];
            crossReferenceRow.Value = "REDCap ID";
            crossReferenceRow.ColumnWidth = 14;

            crossReferenceRow.Offset[0, 1].Value = "ART SPARK ID";
            crossReferenceRow.Offset[0, 1].ColumnWidth = 14;

            crossReferenceRow.Offset[0, 2].Value = "REDCap Text";
            crossReferenceRow.Offset[0, 2].ColumnWidth = 75;

            crossReferenceRow.Offset[0, 3].Value = "ART SPARK Text";
            crossReferenceRow.Offset[0, 3].ColumnWidth = 75;

            crossReferenceRow.Font.Bold = true;

            crossReferenceRow = crossReferenceRow.Offset[1, 0];
            crossReferenceSheet.Select();
            application.ActiveWindow.SplitRow = 1;
            application.ActiveWindow.FreezePanes = true;
        }

        private void BuildRegex()
        {
            chunkExtractor = new Regex(chunkPattern);
            idExtractor = new Regex(idPattern);
            //idExtractorAlt = new Regex(idPatternAlt);
            providerExtractor = new Regex(providerPattern);
            replyExtractor = new Regex(replyPattern);
        }

        private string GetREDCapID(string redcapIdFull)
        {
            string redcapID = string.Empty;
            Match idMatch = idExtractor.Match(redcapIdFull);

            if (idMatch.Success)
            {
                redcapID = idMatch.Groups["id"].Value;
            }
            else
            {
                throw new FormatException("Unable to parse REDCap ID string.");
            }

            return redcapID;
        }

        private List<string> GetReplyPieces(string redcapText)
        {
            List<string> replyPieces = new List<string>();

            // Get everything AFTER "Physician reply:".
            Match providerMatch = providerExtractor.Match(redcapText);

            if (providerMatch.Success)
            {
                string provider_reply = providerMatch.Groups["provider_reply"].Value;

                if (!string.IsNullOrEmpty(provider_reply))
                {
                    foreach (Match match in replyExtractor.Matches(provider_reply))
                    {
                        if (match.Success)
                        {
                            string piece = match.Groups["reply"].Value;

                            if (!string.IsNullOrEmpty(piece))
                            {
                                replyPieces.Add(piece);
                            }
                        }
                    }
                }
            }

            return replyPieces;
        }

        internal void Match()
        {
            application = Globals.ThisAddIn.Application;
            BuildCrossReferenceSheet();

            using (MatchTextForm form = new MatchTextForm())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    redcapIdColumn = form.redcapIdColumn;
                    redcapTextColumn = form.redcapMessageColumn;
                    Worksheet redcapSheet = redcapTextColumn.Worksheet;

                    artSparkIdColumn = form.artIdColumn;
                    artSparkTextColumn = form.artMessageColumn;
                    Worksheet artSparkSheet = artSparkTextColumn.Worksheet;

                    MatchColumns();
                }
            }
        }

        // Run down the REDCap column, extracting the reply & finding its best match in the ART SPARK column.
        private void MatchColumns()
        {
            BuildRegex();
            redcapMessagesMatched = new List<string>();

            int numREDCapRows = Utilities.FindLastRow(redcapIdColumn);

            int rowOffset = 0;
            string redcapID = string.Empty;
            string redcapIdFull = string.Empty;
            string redcapText = string.Empty;

            while (true)
            {
                rowOffset++;

                try
                {
                    redcapText = redcapTextColumn.Offset[rowOffset, 0].Value;

                    // These are the only lines with message text.
                    if (redcapText.StartsWith("<div class"))
                    {
                        redcapID = GetREDCapID(redcapIdColumn.Offset[rowOffset, 0].Value);

                        // If we've already seen this one, skip it.
                        if (!redcapMessagesMatched.Contains(redcapID))
                        {
                            List<string> replyPieces = GetReplyPieces(redcapText);
                            MatchThisText(redcapID, replyPieces);
                        }

                        ReportProgress(rowOffset, numREDCapRows);
                    }
                }
                catch (System.NullReferenceException)
                {
                    break;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    break;
                }
            }

            application.StatusBar = "Complete";
        }

        private void MatchThisText(string redcapID, List<string> replyPieces)
        {
            if (!TooSmall(replyPieces))
            {
                int rowOffset = 0;
                string artSparkID = string.Empty;
                string artSparkText = string.Empty;

                // Find the given text in the ART SPARK column.
                while (true)
                {
                    rowOffset++;

                    try
                    {
                        artSparkID = artSparkIdColumn.Offset[rowOffset, 0].Value.ToString();
                        artSparkText = artSparkTextColumn.Offset[rowOffset, 0].Value;

                        // Are each of the reply pieces present in the original ART SPARK text?
                        if (replyPieces.All(piece => artSparkText.Contains(piece)))
                        {
                            crossReferenceRow.Value = redcapID;
                            crossReferenceRow.Offset[0, 1].Value = artSparkID;
                            crossReferenceRow.Offset[0, 2].Value = string.Join(", ", replyPieces);
                            crossReferenceRow.Offset[0, 3].Value = artSparkText;
                            crossReferenceRow = crossReferenceRow.Offset[1, 0];

                            // No need to match this one if we see it again.
                            redcapMessagesMatched.Add(redcapID);
                            break;
                        }
                    }
                    catch (System.NullReferenceException)
                    {
                        break;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        break;
                    }
                }
            }
        }

        private void ReportProgress(int rowOffset, int numRows)
        {
            // Don't set ScrollRow to where we're writing now or the whole sheet will appear blank.
            // Scroll to a few rows back (but not < 1).
            int topOfSheet = System.Math.Max(1, crossReferenceRow.Row - 20);
            application.ActiveWindow.ScrollRow = topOfSheet;
            application.StatusBar = "Processing row " + rowOffset.ToString() + "/" + numRows.ToString();
        }

        private bool TooSmall(List<string> pieces)
        {
            if (pieces.Count == 0)
            {
                return true;
            }

            string piecesAsembled = string.Join("", pieces);

            // Don't mind small pieces if they add up to something.
            if (piecesAsembled.Length < MIN_SIZE)
            {
                return true;
            }

            return false;
        }
    }
}
