using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class MessageUnpeeler
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private int lastRow;
        private Range selectedColumnRng;
        private const string namePattern = @"From:(?<name>[\w,\s]+)\s+Sent:";
        private Regex nameRegex;

        internal MessageUnpeeler()
        {
            application = Globals.ThisAddIn.Application;
        }

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application, lastRow);

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
/*
        private List<string> GetNames(string line)
        {
            List<string> names = new List<string>();

            foreach (Match match in nameRegex.Matches(line))
            {
                names.Add(match.Groups["name"].Value.ToString());
            }

            return names.Distinct<string>().ToList();
        }
*/

        internal void Scan(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            // Instantiate reusable Regexes.
            nameRegex = new Regex(namePattern);

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                string newColumnName = selectedColumnName + " (Extracted)";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                 newColumnName: newColumnName,
                                                                 side: InsertSide.Right);

                string sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;

                    string[] lines = sourceData.Split(new string[] { "----- " }, StringSplitOptions.None);

                    foreach(string line in lines)
                    {
                        // Grab the -LAST- line that does not start with Message or From:
                        if (!line.Contains("Message") && !line.Contains("From:"))
                        {
                            target.Value = line;
                        }
                    }
                }
            }
        }
    }
}
