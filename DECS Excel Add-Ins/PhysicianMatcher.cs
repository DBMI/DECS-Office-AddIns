using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class PhysicianMatcher
    {
        private Dictionary<string, string> recordIds;
        private List<string> sourceNames;
        private bool quit = false;

        internal PhysicianMatcher()
        {
            recordIds = new Dictionary<string, string>();
        }

        private string AskUserToMatch(string desiredName, List<string> possibleMatches = null)
        {
            string selectedName = string.Empty;

            using (NameMatchForm form = new NameMatchForm(desiredName, sourceNames, possibleMatches))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.selectedName;
                }

                if (result == DialogResult.Abort)
                {
                    quit = true;
                }
            }

            return selectedName;
        }

        private void BuildIdDictionary(Range sourceColumn, Range idColumn)
        {
            recordIds.Clear();
            int iRowOffset = 0;
            sourceNames = new List<string>();

            while (true)
            {
                try
                {
                    iRowOffset++;
                    string sourceName = sourceColumn.Offset[iRowOffset].Value2.ToString();

                    if (!recordIds.ContainsKey(sourceName))
                    {
                        try
                        {
                            string idString = idColumn.Offset[iRowOffset].Value2.ToString();
                            recordIds.Add(sourceName, idString);

                            // Record the ORIGINAL list of names
                            // (before we start adding matches from target list).
                            sourceNames.Add(sourceName);
                        }
                        // If there's no ID, skip to next row.
                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                        {
                            continue;
                        }
                    }
                }
                // If there's no next row, we're done.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }
            }
        }

        private Worksheet FindOrCreateTranslationTable(Range sourceColumn, Range idColumn, Range targetColumn)
        {
            Worksheet translationTable = null;
            Workbook workbook = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;

            try
            {
                translationTable = workbook.Worksheets["Name Translation Table"];
                Dictionary<string, Range> translationTableColumns = Utilities.GetColumnRangeDictionary(translationTable);
                Range translationTableNamesColumn = translationTableColumns["Unique Target Names"];
                Range translationTableIdsColumn = translationTableColumns["Id of Matching Name"];
                BuildIdDictionary(translationTableNamesColumn, translationTableIdsColumn);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                BuildIdDictionary(sourceColumn, idColumn);
                translationTable = Utilities.CreateNewNamedSheet("Name Translation Table");
                Range uniqueTargetNamesColumn = (Range)translationTable.Cells[1, 1];
                uniqueTargetNamesColumn.Offset[0, 0].Value = "Unique Target Names";
                List<string> uniqueTargetNames = Utilities.ExtractColumnUnique(targetColumn);

                for (int i = 1; i <= uniqueTargetNames.Count; i++)
                {
                    uniqueTargetNamesColumn.Offset[i, 0].Value = uniqueTargetNames[i - 1];
                }

                InsertMatchingNameAndId(uniqueTargetNamesColumn);
            }

            return translationTable;
        }

        private void InsertIdWhereNamesMatch(Range tgtColumn)
        {
            Range newColumn = Utilities.InsertNewColumn(tgtColumn, "Record ID");
            int numRows = Utilities.FindLastRow(tgtColumn);
            string idString;
            int iRowOffset = 0;

            while (true)
            {
                try
                {
                    iRowOffset += 1;
                    string thisName = tgtColumn.Offset[iRowOffset].Value2.ToString();

                    // Exact match?
                    if (recordIds.ContainsKey(thisName))
                    {
                        idString = recordIds[thisName];
                        newColumn.Offset[iRowOffset].Value = idString;
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }
            }
        }

        private void InsertMatchingNameAndId(Range tgtColumn)
        {
            Range matchingNameColumn = Utilities.InsertNewColumn(tgtColumn, "Matching Name");
            Range matchDetailsColumn = Utilities.InsertNewColumn(matchingNameColumn, "Match Details");
            Range matchingIdColumn = Utilities.InsertNewColumn(matchDetailsColumn, "Id of Matching Name");
            string idString;
            int iRowOffset = 0;

            while (true)
            {
                try
                {
                    iRowOffset += 1;
                    string thisName = tgtColumn.Offset[iRowOffset].Value2.ToString();

                    // Exact match?
                    if (recordIds.ContainsKey(thisName))
                    {
                        matchingNameColumn.Offset[iRowOffset].Value = thisName;
                        matchDetailsColumn.Offset[iRowOffset].Value = "Exact";
                        idString = recordIds[thisName];
                        matchingIdColumn.Offset[iRowOffset].Value = idString;
                    }
                    else
                    {
                        List<string> allNames = recordIds.Keys.ToList();
                        NameMatch nameMatch = Utilities.FindClosestMatch(allNames, thisName, maxDistanceAllowed: 0.25);

                        if (nameMatch.IsMatch())
                        {
                            matchingNameColumn.Offset[iRowOffset].Value = nameMatch.BestMatch();
                            matchDetailsColumn.Offset[iRowOffset].Value = nameMatch.MatchType();
                            idString = recordIds[nameMatch.BestMatch()];
                            matchingIdColumn.Offset[iRowOffset].Value = idString;

                            // Put target name into dictionary so it's easier to find next time.
                            recordIds[thisName] = idString;
                        }
                        else
                        {
                            // Here we don't want to compare against the ever-growing dictionary keys,
                            // but against the original names associated with the record IDs.
                            List<string> possibleMatches = Utilities.MightMatch(sourceNames, thisName);

                            if (possibleMatches.Count > 0)
                            {
                                string userSelection = AskUserToMatch(desiredName: thisName,
                                                                      possibleMatches: possibleMatches);

                                // If user has pressed Quit, stop asking.
                                if (quit)
                                {
                                    break;
                                }

                                if (string.IsNullOrEmpty(userSelection))
                                {
                                    // Put target name into dictionary so we STOP asking the user.
                                    recordIds[thisName] = string.Empty;
                                }
                                else
                                {
                                    matchingNameColumn.Offset[iRowOffset].Value = userSelection;
                                    matchDetailsColumn[iRowOffset].Value = TypeOfMatch.UserSelected.ToString();
                                    idString = recordIds[userSelection];
                                    matchingIdColumn.Offset[iRowOffset].Value = idString;

                                    // Put target name into dictionary so it's easier to find next time.
                                    recordIds[thisName] = idString;
                                }
                            }
                        }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }
            }
        }

        internal void Match()
        {
            using (MatchSetupForm form = new MatchSetupForm())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    Range sourceColumn = form.sourceColumn;
                    Range idColumn = form.idColumn;
                    Range targetColumn = form.targetColumn;

                    FindOrCreateTranslationTable(sourceColumn, idColumn, targetColumn);

                    if (!quit)
                    {
                        InsertIdWhereNamesMatch(targetColumn);
                    }
                }
            }
        }
    }
}
