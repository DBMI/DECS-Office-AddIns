using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Breaks an Excel spreadsheet into groups based on a selected column.
     */
    internal class Striper
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private Range selectedColumnRng;
        private Dictionary<string, Block> sourceBlocks;
        private Worksheet thisWorksheet;
        private XlRgbColor gray = XlRgbColor.rgbLightGray;

        internal Striper()
        {
            application = Globals.ThisAddIn.Application;
        }

        private bool FindSelectedCategory(Worksheet worksheet)
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
                        string selectedColumnName = form.selectedColumns[0];
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

        internal void Run(Worksheet worksheet)
        {
            thisWorksheet = worksheet;
            
            // Remove all old shading.
            Utilities.ClearWorksheet(worksheet);
            
            if (FindSelectedCategory(worksheet))
            {
                // Figure out the distinct category values & where they are.
                sourceBlocks = Utilities.IdentifyBlocks(selectedColumnRng);
                List<string> blockNames = new List<string>(sourceBlocks.Keys);

                int blockIndex = 0;

                // Process each block in order.
                foreach (string blockName in blockNames)
                {
                    // Stripe the even numbered blocks.
                    if (blockIndex % 2 == 0)
                    {
                        // Find the Block of rows in the sheet corresponding to this value.
                        Block thisBlock = sourceBlocks[blockName];
                        thisBlock.shade(thisWorksheet, gray);
                    }

                    blockIndex++;
                }
            }
        }
    }
}
