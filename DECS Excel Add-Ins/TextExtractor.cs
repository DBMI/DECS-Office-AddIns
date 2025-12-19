using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class TextExtractor
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private Range selectedColumnRng;

        private Regex patientMessageExtractor;
        private const string PATIENT_MESSAGE_PATTERN = @"Patient message:<\/strong><br><br \/><span style=""font-weight: normal;"">(?<patientMessage>.*)<\/span><\/p><p>";
        private Regex providerReplyExtractor;
        private const string PROVIDER_REPLY_PATTERN = @"Provider reply:<\/strong><br><br \/><span style=""font-weight: normal;"">(?<providerReply>.*)<\/span><\/p><\/div>";

        internal TextExtractor()
        {
            application = Globals.ThisAddIn.Application;
            patientMessageExtractor = new Regex(PATIENT_MESSAGE_PATTERN);
            providerReplyExtractor = new Regex(PROVIDER_REPLY_PATTERN);
        }

        internal void Extract(Worksheet worksheet)
        {

            if (FindSelectedColumn(worksheet))
            {
                // Make room for TWO new columns.
                Range patientMessageColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                 newColumnName: "Patient message",
                                                                 side: InsertSide.Right);
                Range providerReplyColumn = Utilities.InsertNewColumn(range: patientMessageColumn,
                                                                 newColumnName: "Provider reply",
                                                                 side: InsertSide.Right);

                string sourceData;
                Range patientMessageTarget;
                Range providerReplyTarget;
                int rowNumber = 1;

                while (true)
                {
                    rowNumber++;
                    patientMessageTarget = (Range)worksheet.Cells[rowNumber, patientMessageColumn.Column];
                    providerReplyTarget = (Range)worksheet.Cells[rowNumber, providerReplyColumn.Column];

                    try
                    {
                        sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                        
                        // Get everything AFTER "Patient message:".
                        Match patientMessageMatch = patientMessageExtractor.Match(sourceData);

                        if (patientMessageMatch.Success) 
                        {
                            string patientMessage = patientMessageMatch.Groups["patientMessage"].Value;

                            if (!string.IsNullOrEmpty(patientMessage))
                            {
                                patientMessageTarget.Value = patientMessage;
                            }
                        }

                        // Get everything AFTER "Physician reply:".
                        Match providerMatch = providerReplyExtractor.Match(sourceData);

                        if (providerMatch.Success)
                        {
                            string providerReply = providerMatch.Groups["providerReply"].Value;

                            if (!string.IsNullOrEmpty(providerReply))
                            {
                                providerReplyTarget.Value = providerReply;
                            }
                        }
                    }
                    catch (System.ArgumentNullException)
                    {
                        break;
                    }
                }
            }
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


    }
}
