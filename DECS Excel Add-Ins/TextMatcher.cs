using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
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

        private char ASCII_QUOTE = (char)39;
        private char UNICODE_QUOTE = (char)8217;

        private string COPYRIGHT = @"©";
        private string REDCap_MANGLED_COPYRIGHT = @"Â©";
        private string REDCap_MANGLED_DASH = @"â€""";
        private string REDCap_MANGLED_LEFT_QUOTES = @"â€œ";
        private string REDCap_MANGLED_RIGHT_QUOTES = @"â€";
        private string UNICODE_DASH = "–";
        private string UNICODE_LEFT_QUOTES = "“";
        private string UNICODE_RIGHT_QUOTES = "”";


        private Regex chunkExtractor;
        private const string CHUNK_PATTERN = @"(\>[^\<\>]{3,}\<)+";
        private Regex idExtractor;
        private const string ID_PATTERN = @"(?:icr)*(?:nc)*(?:message)*(?<id>[\d\w]*)(?:message)*";
        private Regex providerExtractor;
        private const string PROVIDER_PATTERN = @"Provider reply:<\/strong><br><br \/><span style=""font-weight: normal;"">(?<providerReply>.*)<\/span><\/p><\/div>";
        private const string REDACTED = "Redacted";
        private Regex replyExtractor;
        private const string REPLY_PATTERN = @"(?<reply>(\d|\s|\w|,|\.|\<\d|(?<!\w)\>|\/|\(|\)|\[|\]|'|""|\?|!|:|;|\*|\-|©|“|”|–)*)";

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

            crossReferenceRow.Offset[0, 2].Value = "# ART SPARK Msgs matched";
            crossReferenceRow.Offset[0, 2].ColumnWidth = 14;

            crossReferenceRow.Offset[0, 3].Value = "REDCap Text";
            crossReferenceRow.Offset[0, 3].ColumnWidth = 75;

            crossReferenceRow.Offset[0, 4].Value = "ART SPARK Text";
            crossReferenceRow.Offset[0, 4].ColumnWidth = 75;

            crossReferenceRow.Font.Bold = true;

            crossReferenceRow = crossReferenceRow.Offset[1, 0];
            crossReferenceSheet.Select();
            application.ActiveWindow.SplitRow = 1;
            application.ActiveWindow.FreezePanes = true;
        }

        private void BuildRegex()
        {
            chunkExtractor = new Regex(CHUNK_PATTERN);
            idExtractor = new Regex(ID_PATTERN);
            providerExtractor = new Regex(PROVIDER_PATTERN);
            replyExtractor = new Regex(REPLY_PATTERN);
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
                string providerReply = providerMatch.Groups["providerReply"].Value;

                if (!string.IsNullOrEmpty(providerReply))
                {
                    foreach (Match match in replyExtractor.Matches(providerReply))
                    {
                        if (match.Success)
                        {
                            string piece = match.Groups["reply"].Value;

                            if (!string.IsNullOrEmpty(piece) && !string.Equals(piece, REDACTED))
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
                        // REDCap apparently injects an additional char before the © symbol.
                        // Fix that here or it will break matches.
                        redcapText = redcapText.Replace(REDCap_MANGLED_COPYRIGHT, COPYRIGHT);
                        redcapText = redcapText.Replace(REDCap_MANGLED_DASH, UNICODE_DASH);
                        redcapText = redcapText.Replace(REDCap_MANGLED_LEFT_QUOTES, UNICODE_LEFT_QUOTES);
                        redcapText = redcapText.Replace(REDCap_MANGLED_RIGHT_QUOTES, UNICODE_RIGHT_QUOTES);

                        redcapID = GetREDCapID(redcapIdColumn.Offset[rowOffset, 0].Value);

                        // If we've already seen this one, skip it.
                        if (redcapMessagesMatched.Contains(redcapID))
                        {
                            application.StatusBar = "Skipping " + redcapID;
                        }
                        else
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
                int numMatchesForThisREDCapMsg = 0;
                List<string> artSparkIDsSeen = new List<string>();

                // Find the given text in the ART SPARK column.
                while (true)
                {
                    rowOffset++;
                    crossReferenceRow.Offset[0, 2].Value = numMatchesForThisREDCapMsg.ToString();

                    try
                    {
                        // Even if no matches, show what we TRIED to match.
                        crossReferenceRow.Value = redcapID;
                        crossReferenceRow.Offset[0, 3].Value = string.Join(", ", replyPieces);

                        artSparkID = artSparkIdColumn.Offset[rowOffset, 0].Value.ToString();

                        // If we've already tested this ART SPARK message, no need to test it again.
                        if (artSparkIDsSeen.Contains(artSparkID))
                        {
                            continue;
                        }

                        artSparkIDsSeen.Add(artSparkID);
                        artSparkText = artSparkTextColumn.Offset[rowOffset, 0].Value;

                        // Don't get confused by fancy UNICODE quotes (which apparently get translated in REDCap extraction.)
                        artSparkText = artSparkText.Replace(UNICODE_QUOTE, ASCII_QUOTE);

                        // Are each of the reply pieces present in the original ART SPARK text?
                        if (replyPieces.All(piece => artSparkText.Contains(piece)))
                        {
                            numMatchesForThisREDCapMsg++;
                            crossReferenceRow.Value = redcapID;
                            crossReferenceRow.Offset[0, 1].Value = artSparkID;
                            crossReferenceRow.Offset[0, 2].Value = numMatchesForThisREDCapMsg.ToString();
                            crossReferenceRow.Offset[0, 3].Value = string.Join(", ", replyPieces);
                            crossReferenceRow.Offset[0, 4].Value = artSparkText;
                            crossReferenceRow = crossReferenceRow.Offset[1, 0];

                            // No need to match this one if we see it again.
                            redcapMessagesMatched.Add(redcapID);
                        }
                    }
                    catch (System.NullReferenceException)
                    {
                        break;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        if (numMatchesForThisREDCapMsg == 0)
                        {
                            // Move to the next row.
                            crossReferenceRow = crossReferenceRow.Offset[1, 0];
                        }

                        break;
                    }
                }               
            }
        }

        private void ReportProgress(int rowOffset, int numRows)
        {
            // Don't set ScrollRow to where we're writing now or the whole sheet will appear blank.
            // Scroll to a few rows back (but not < 1).
            int topOfSheet = Math.Max(1, crossReferenceRow.Row - 20);
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
