using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using TextBox = System.Windows.Forms.TextBox;
using ToolTip = System.Windows.Forms.ToolTip;

namespace DECS_Excel_Add_Ins
{
    internal class Utilities
    {
        internal static Range AllAvailableRows(Worksheet sheet)
        {
            Range firstCell = (Range)sheet.Cells[1, 1];
            Range lastCell = (Range)sheet.Cells[Utilities.FindLastRow(sheet), 1];
            Range allRows = (Range)sheet.Range[firstCell, lastCell];
            return allRows;
        }

        internal static void ClearRegexInvalid(TextBox textBox)
        {
            if (textBox == null)
                return;

            // Clear any previous highlighting.
            textBox.BackColor = Color.White;

            // Remove the MouseHover EventHandler.
            DetachEvents(textBox);
        }

        // Convert Excel-formatted date to SQL style.
        internal static string ConvertExcelDate(string cellContents)
        {
            string convertedContents = null;

            try
            {
                double d = double.Parse(cellContents);
                DateTime conv = DateTime.FromOADate(d);
                convertedContents = conv.ToString("yyyy-MM-dd");
            }
            catch (System.FormatException)
            {
                // Probably trying to convert the name "Date" to a Double in order to create DateTime object.
            }

            return convertedContents;
        }

        internal static int CountCellsWithData(Range rng, int lastRow)
        {
            int numCellsWithData = 0;
            Range thisCell;

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
            {
                thisCell = rng.Cells[rowNumber];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    numCellsWithData = rowNumber;
                    break;
                }
            }

            return numCellsWithData;
        }

        public static void DetachEvents(TextBox textBox)
        {
            object objNew = textBox
                .GetType()
                .GetConstructor(new Type[] { })
                .Invoke(new object[] { });
            PropertyInfo propEvents = textBox
                .GetType()
                .GetProperty("Events", BindingFlags.NonPublic | BindingFlags.Instance);

            EventHandlerList eventHandlerList_obj = (EventHandlerList)
                propEvents.GetValue(textBox, null);
            eventHandlerList_obj.Dispose();
        }

        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Columns, including cells that contains formulas that result in blank values
            return sheet.UsedRange.Columns.Count;
        }

        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastRow(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Row, including cells that contains formulas that result in blank values
            return sheet.UsedRange.Rows.Count;
        }

        internal static ToolTip FindToolTip(TextBox textBox)
        {
            ToolTip toolTip = null;
            Form form = textBox.FindForm();

            if (form != null)
            {
                // https://stackoverflow.com/a/42113517/18749636
                Type typeForm = form.GetType();
                FieldInfo fieldInfo = typeForm.GetField(
                    "components",
                    BindingFlags.Instance | BindingFlags.NonPublic
                );
                IContainer parent = (IContainer)fieldInfo.GetValue(form);
                List<ToolTip> ToolTipList = parent.Components.OfType<ToolTip>().ToList();

                if (ToolTipList.Count > 0)
                {
                    toolTip = ToolTipList[0];
                }
            }

            return toolTip;
        }

        // Turn the scope-of-work filename into a .sql filename.
        internal static string FormOutputFilename(
            string filename,
            string filenameAddon = "",
            string filetype = ".sql",
            bool shortVersion = false
        )
        {
            string dir = Path.GetDirectoryName(filename);
            string justTheFilename = Path.GetFileNameWithoutExtension(filename) + filenameAddon;
            string sqlFilename = Path.Combine(dir, justTheFilename + filetype);

            if (shortVersion)
            {
                sqlFilename = justTheFilename + filetype;
            }

            return sqlFilename;
        }

        internal static List<string> GetColumnNames(Worksheet sheet)
        {
            List<string> names = new List<string>();
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                names.Add(range.Value.ToString());

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return names;
        }

        internal static List<string> GetColumnNames(List<Range> selectedColumns)
        {
            List<string> names = new List<string>();

            if (selectedColumns != null && selectedColumns.Count > 0)
            {
                // Search along row 1.
                foreach (Range col in selectedColumns)
                {
                    Range topCell = col.Cells[1, 1];

                    try
                    {
                        names.Add(topCell.Value.ToString());
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        continue;
                    }
                }
            }

            return names;
        }

        // https://stackoverflow.com/q/21219797/18749636
        internal static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        internal static bool HasData(Range rng, int lastRow)
        {
            bool hasData = false;
            Range thisCell;

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
            {
                thisCell = rng.Cells[rowNumber];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                    hasData = true;
                    break;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            }

            return hasData;
        }

        internal static Range InsertNewColumn(Range range, string newColumnName)
        {
            int columnNumber = range.Column;
            Worksheet sheet = range.Worksheet;
            sheet.Columns[columnNumber].Insert();

            Range newRange = range.Offset[0, -1];
            newRange.Value2 = newColumnName;
            return newRange;
        }

        internal static RuleValidationResult IsRegexValid(string regexText)
        {
            // Empty strings are not errors.
            if (!string.IsNullOrEmpty(regexText))
            {
                try
                {
                    Regex regex = new Regex(regexText);
                }
                catch (ArgumentException ex)
                {
                    return new RuleValidationResult(ex);
                }
            }

            return new RuleValidationResult();
        }

        internal static void MarkRegexInvalid(TextBox textBox, string message)
        {
            if (textBox == null)
                return;

            // Highlight box to show RegEx is invalid.
            textBox.BackColor = Color.Pink;

            ToolTip toolTip = FindToolTip(textBox);

            if (toolTip != null)
            {
                Action<object, System.EventArgs> mouseHover = (sender, e) =>
                {
                    toolTip.SetToolTip(textBox, message);
                };

                textBox.MouseHover += new System.EventHandler(mouseHover);
            }
        }

        // Open the output StreamWriter object, understanding that we might
        // have to substitute a shorter version of the output filename
        // if the default filename is too long.
        internal static (StreamWriter writer, string openedFilename) OpenOutput(
            string inputFilename,
            string filenameAddon = "",
            string filetype = ".sql"
        )
        {
            string outputFilename = Utilities.FormOutputFilename(
                filename: inputFilename,
                filetype: filetype,
                filenameAddon: filenameAddon,
                shortVersion: false
            );
            StreamWriter writer_obj;

            try
            {
                writer_obj = new StreamWriter(outputFilename);
            }
            // https://stackoverflow.com/a/19329123/18749636
            catch (Exception ex)
                when (ex is System.IO.PathTooLongException || ex is System.NotSupportedException)
            {
                outputFilename = Utilities.FormOutputFilename(
                    filename: inputFilename,
                    filetype: filetype,
                    shortVersion: true
                );
                writer_obj = new StreamWriter(outputFilename);
            }

            return (writer: writer_obj, openedFilename: outputFilename);
        }

        internal static Range TopOfNamedColumn(Worksheet sheet, string columnName)
        {
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                if (range.Value == columnName)
                {
                    return range;
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return null;
        }

        internal static void WarnColumnNotFound(string columnName)
        {
            string message = "Column '" + columnName + "' not found.";
            string title = "Not Found";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);
        }
    }
}
