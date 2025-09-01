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

        internal PhysicianMatcher()
        {
            recordIds = new Dictionary<string, string>();
        }

        private string AskUserToMatch(string desiredName)
        {
            string selectedName = string.Empty;
            List<string> possibleNames = recordIds.Keys.ToList();

            using (NameMatchForm form = new NameMatchForm(desiredName, possibleNames))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.selectedName;
                }
            }

            return selectedName;
        }

        private void BuildIdDictionary(Range sourceColumn, Range idColumn)
        {
            recordIds.Clear();
            int iRowOffset = 0;

            while (true)
            {
                try
                {
                    iRowOffset++;
                    string sourceName = sourceColumn.Offset[iRowOffset].Value2.ToString();

                    if (!recordIds.ContainsKey(sourceName))
                    {
                        string idString = idColumn.Offset[iRowOffset].Value2.ToString();
                        recordIds.Add(sourceName, idString);
                    }
                }
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
                    else
                    {
                        List<string> allNames = recordIds.Keys.ToList();
                        string closestName = Utilities.FindClosestMatch(allNames, thisName, maxDistanceAllowed: 0.25);

                        if (string.IsNullOrEmpty(closestName))
                        {
                            string userSelection = AskUserToMatch(thisName);

                            if (string.IsNullOrEmpty(userSelection))
                            {
                                // Put target name into dictionary so we STOP asking the user.
                                recordIds[thisName] = string.Empty;
                            }
                            else
                            {
                                idString = recordIds[userSelection];
                                newColumn.Offset[iRowOffset].Value = idString;

                                // Put target name into dictionary so it's easier to find next time.
                                recordIds[thisName] = idString;
                            }
                        }
                        else
                        {
                            idString = recordIds[closestName];
                            newColumn.Offset[iRowOffset].Value = idString;

                            // Put target name into dictionary so it's easier to find next time.
                            recordIds[thisName] = idString;
                        }
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
            Range matchingIdColumn = Utilities.InsertNewColumn(matchingNameColumn, "Id of Matching Name");
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
                        idString = recordIds[thisName];
                        matchingIdColumn.Offset[iRowOffset].Value = idString;
                    }
                    else
                    {
                        List<string> allNames = recordIds.Keys.ToList();
                        string closestName = Utilities.FindClosestMatch(allNames, thisName, maxDistanceAllowed: 0.25);

                        if (string.IsNullOrEmpty(closestName))
                        {
                            string userSelection = AskUserToMatch(thisName);

                            if (string.IsNullOrEmpty(userSelection))
                            {
                                // Put target name into dictionary so we STOP asking the user.
                                recordIds[thisName] = string.Empty;
                            }
                            else
                            {
                                matchingNameColumn.Offset[iRowOffset].Value = userSelection;
                                idString = recordIds[userSelection];
                                matchingIdColumn.Offset[iRowOffset].Value = idString;

                                // Put target name into dictionary so it's easier to find next time.
                                recordIds[thisName] = idString;
                            }
                        }
                        else
                        {
                            matchingNameColumn.Offset[iRowOffset].Value = closestName;
                            idString = recordIds[closestName];
                            matchingIdColumn.Offset[iRowOffset].Value = idString;

                            // Put target name into dictionary so it's easier to find next time.
                            recordIds[thisName] = idString;
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

                    Worksheet translationTable = FindOrCreateTranslationTable(sourceColumn, idColumn, targetColumn);
                    InsertIdWhereNamesMatch(targetColumn);
                }
            }
        }
    }
}
