using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    public enum MessageDirection
    {
        FromPatient,
        ToPatient,
        None
    }

    internal class MessageUnpeeler
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private readonly string[] DELIMITERS = { "----- ", "          ", "Subject:" };
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

        private MessageDirection ParseDirectionFromColumnName(string columnName)
        {
            MessageDirection messageDirection = MessageDirection.None;

            if (!string.IsNullOrEmpty(columnName)) 
            { 
                if (columnName.ToLower().Contains("from patient"))
                {
                    messageDirection = MessageDirection.FromPatient;
                }
                else if (columnName.ToLower().Contains("to patient"))
                {
                    messageDirection = MessageDirection.ToPatient;
                }
                else
                {
                    // If we can't figure it out from the column name, ask user directly;
                    using (MessageDirectionForm form = new MessageDirectionForm())
                    {
                        var result = form.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            messageDirection = form.direction;
                        }
                    }
                }
            }

            return messageDirection;
        }

        internal void Scan(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            // Instantiate reusable Regexes.
            nameRegex = new Regex(namePattern);

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                MessageDirection messageDirection = ParseDirectionFromColumnName(selectedColumnName);
                
                // If we can't decipher the message direction, quit.
                if (messageDirection == MessageDirection.None)
                {
                    return;
                }

                string newColumnName = selectedColumnName + " (Extracted)";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                 newColumnName: newColumnName,
                                                                 side: InsertSide.Right);

                string sourceData;
                string targetData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value.ToString();

                    string[] lines = sourceData.Split(DELIMITERS, StringSplitOptions.None);

                    // In case the message doesn't contain one of the delimiters,
                    // initialize with the raw message.
                    targetData = sourceData;

                    foreach (string line in lines)
                    {
                        // Skip empty lines.
                        if (line.Trim().Length > 0)
                        {
                            // Find lines that do not contain Message or From:or MyChart boilerplate.
                            if (!line.Contains("MyChart Guidelines:") &&
                                !line.Contains("Message") &&
                                !line.Contains("From:"))
                            {
                                targetData = line.Trim();

                                // If message is TO the patient:
                                // Grab the -FIRST- such line, so we're done.
                                if (messageDirection == MessageDirection.ToPatient)
                                {
                                    break;
                                }
                                // If message is FROM the patient:
                                // Grab the -LAST- such line, so keep parsing.
                            }
                        }
                    }

                    target.Value2 = targetData;
                }
            }
        }
    }
}
