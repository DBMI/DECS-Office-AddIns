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
            if (textBox == null) return;

            // Clear any previous highlighting.
            textBox.BackColor = Color.White;

            // Remove the MouseHover EventHandler.
            DetachEvents(textBox);
        }

        public static void DetachEvents(TextBox textBox)
        {
            object objNew = textBox.GetType().GetConstructor(new Type[] { }).Invoke(new object[] { });
            PropertyInfo propEvents = textBox.GetType().GetProperty("Events", BindingFlags.NonPublic | BindingFlags.Instance);

            EventHandlerList eventHandlerList_obj = (EventHandlerList)propEvents.GetValue(textBox, null);
            eventHandlerList_obj.Dispose();
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
                FieldInfo fieldInfo = typeForm.GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
                IContainer parent = (IContainer)fieldInfo.GetValue(form);
                List<ToolTip> ToolTipList = parent.Components.OfType<ToolTip>().ToList();

                if (ToolTipList.Count > 0)
                {
                    toolTip = ToolTipList[0];
                }
            }

            return toolTip;
        }

        // Turn the statement of work filename into a .sql filename.
        internal static string FormOutputFilename(string filename, string filetype = ".sql", bool short_version = false)
        {
            string dir = Path.GetDirectoryName(filename);
            string just_the_filename = Path.GetFileNameWithoutExtension(filename);
            string sql_filename = Path.Combine(dir, just_the_filename + filetype);

            if (short_version)
            {
                sql_filename = just_the_filename + filetype;
            }

            return sql_filename;
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

        // https://stackoverflow.com/q/21219797/18749636
        internal static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        internal static Range InsertnewColumn(Range range, string newColumnName)
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
            if (textBox == null) return;

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
        internal static (StreamWriter writer, string opened_filename) OpenOutput(string input_filename, string filetype = ".sql")
        {
            string output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: false);
            StreamWriter writer_obj;

            try
            {
                writer_obj = new StreamWriter(output_filename);
            }
            // https://stackoverflow.com/a/19329123/18749636
            catch (Exception ex) when (
                ex is System.IO.PathTooLongException
                || ex is System.NotSupportedException)
            {
                output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: true);
                writer_obj = new StreamWriter(output_filename);
            }

            return (writer: writer_obj, opened_filename: output_filename);
        }

        // Reassure the user that we've created the desired output file,
        // and display the file once they've seen the message.
        internal static void ShowResults(string output_filename)
        {
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            string message = "Created file '" + output_filename + "'.";
            DialogResult result = MessageBox.Show(message, "Success", buttons);

            if (result == DialogResult.OK)
            {
                Process.Start(output_filename);
            }
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