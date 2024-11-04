using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Breaks an Excel spreadsheet into separate sheets based on a selected column.
     */
    internal class ListChopper
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private BackgroundWorker chopper1;
        private Formatter formatter;
        private int lastRowInSheet;
        private Dictionary<string, Worksheet> newWorksheets;
        private Range selectedColumnRng;
        private Dictionary<string, Block> sourceBlocks;
        private Worksheet thisWorksheet;

        internal ListChopper()
        {
            application = Globals.ThisAddIn.Application;
            formatter = new Formatter();
        }

        private void BuildNewSheet(BackgroundWorker bw, string newName)
        {
            // Look up the existing sheet by name.
            Worksheet newSheet = newWorksheets[newName];

            // Copy over the header to this new sheet.
            Utilities.CopyRow(thisWorksheet, 0, newSheet, 0);

            // Find the Block of rows in the source Worksheet corresponding to this name.
            Block sourceBlock = sourceBlocks[newName];

            // Where to put them in the target sheet.
            Block targetBlock = sourceBlock.sameSize(1);    // start after the header.

            // Copy as a block.
            Utilities.CopyBlock(thisWorksheet, sourceBlock, newSheet, targetBlock);

            // Now that we've populated this new sheet, clean it up & let's look at it.
            formatter.Format(newSheet);
            newSheet.Select();
        }

        // Create the needed sheets in same order in which names are provided.
        private void CreateNeededSheets(List<string> newSheetNames)
        {
            newWorksheets = new Dictionary<string, Worksheet>();

            // Create new sheet for each name.
            foreach (string newName in newSheetNames)
            {
                Worksheet newSheet = Utilities.CreateNewNamedSheet(thisWorksheet, newName);
                newWorksheets.Add(newName, newSheet);
            }
        }

        private bool FindSelectedCategory(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application, lastRowInSheet);

            if (selectedColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string selectedColumnName = form.selectedCategory;
                        selectedColumnRng = Utilities.TopOfNamedColumn(worksheet, selectedColumnName);
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

        private void Chopper1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> newSheetNames = (List<string>)e.Argument;

            // Do not access the form's BackgroundWorker reference directly.
            // Instead, use the reference provided by the sender parameter.
            BackgroundWorker bw = sender as BackgroundWorker;

            // Start the time-consuming operation(s).

            // Initialize new sheet for each name, then populate with matching rows.
            foreach (string newName in newSheetNames)
            {
                BuildNewSheet(bw, newName);
            }

            // If the operation was canceled by the user,
            // set the DoWorkEventArgs.Cancel property to true.
            if (bw.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        internal void Scan(Worksheet worksheet)
        {
            thisWorksheet = worksheet;
            lastRowInSheet = worksheet.UsedRange.Rows.Count;
            
            if (FindSelectedCategory(worksheet))
            {
                string selectedColumnName = Utilities.GetColumnName(selectedColumnRng);

                // Figure out the distinct category values & where they are.
                sourceBlocks = Utilities.IdentifyBlocks(selectedColumnRng, lastRowInSheet);

                // Create new worksheets--one per distinct value.
                List<string> newSheetNames = new List<string>(sourceBlocks.Keys);
                newSheetNames.Sort();
                CreateNeededSheets(newSheetNames);

                // To avoid locking up the main thread, send the copying off to a BackgroundWorker.
                chopper1 = new BackgroundWorker();
                chopper1.DoWork += new DoWorkEventHandler(Chopper1_DoWork);
                chopper1.RunWorkerAsync(newSheetNames);
            }
        }
    }
}
