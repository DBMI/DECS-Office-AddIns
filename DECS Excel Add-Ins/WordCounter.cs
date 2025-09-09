using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class WordCounter
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private int lastRow;
        private Range selectedColumnRng;

        internal WordCounter()
        {
            application = Globals.ThisAddIn.Application;
        }

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application);

            if (selectedColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        selectedColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        internal void Scan(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();

                // Make room for new columns.
                Range wordCountColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                  newColumnName: selectedColumnName + " (# Words)",
                                                                  side: InsertSide.Right);
                Range charCountColumn = Utilities.InsertNewColumn(range: wordCountColumn,
                                                                  newColumnName: selectedColumnName + " (# Char)",
                                                                  side: InsertSide.Right);

                string sourceData;
                Range targetWord;
                Range targetChar;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                    targetWord = (Range)worksheet.Cells[rowNumber, wordCountColumn.Column];
                    targetChar = (Range)worksheet.Cells[rowNumber, charCountColumn.Column];

                    if (sourceData == null)
                    {
                        targetWord.Value2 = 0;
                        targetChar.Value2 = 0;
                    }
                    else
                    {
                        // Count word characters separated by non-word boundaries.
                        int wordCount = Regex.Matches(sourceData, @"\b\w+\b").Count;
                        targetWord.Value2 = wordCount;
                        targetChar.Value2 = sourceData.Length;
                    }
                }
            }
        }
    }
}
