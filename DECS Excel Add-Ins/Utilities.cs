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
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using TextBox = System.Windows.Forms.TextBox;
using ToolTip = System.Windows.Forms.ToolTip;

namespace DECS_Excel_Add_Ins
{
    public enum InsertSide
    {
        Left,
        Right
    }
    /**
     * @brief Useful tools
     */
    internal class Utilities
    {
        /// <summary>
        /// Builds a Range object containing the first column of all rows up to the last row containing any data.
        /// </summary>
        /// <param name="sheet">ActiveWorksheet.</param>
        /// <returns>Range</returns>
        internal static Range AllAvailableRows(Worksheet sheet)
        {
            Range firstCell = (Range)sheet.Cells[1, 1];
            Range lastCell = (Range)sheet.Cells[Utilities.FindLastRow(sheet), 1];
            Range allRows = (Range)sheet.Range[firstCell, lastCell];
            return allRows;
        }

        internal static string CleanColumnNamesForSQL(string columnName)
        {
            // Remove stuff that breaks the SQL script.   
            string niceColumnName = columnName.Trim().
                                        Replace(" ", "_").
                                        Replace("/", "_").
                                        Replace("-", "_").
                                        Replace(",", "").
                                        Replace("(", "").
                                        Replace(")", "").
                                        Replace("__", "_").
                                        Replace("__", "_").
                                        ToUpper();
            return niceColumnName;
        }

        internal static List<string> CleanColumnNamesForSQL(List<string> columnNames)
        {
            List<string> niceColumnNames = new List<string>();

            foreach(string columnName in columnNames)
            {
                niceColumnNames.Add(Utilities.CleanColumnNamesForSQL(columnName));
            }

            return niceColumnNames;
        }


        // Remove quotes that break the SQL import script.
        internal static string CleanDataForSQL(string row)
        {
            if (string.IsNullOrEmpty(row))
            {
                return string.Empty;
            }

            string niceRow = row.Trim();
            int stringLength = niceRow.Length;

            // Remove trailing quotes.
            string pattern = @"([^'])'$";
            string replacement = "$1";
            niceRow = Regex.Replace(niceRow, pattern, replacement);

            // Remove trailing slash.
            pattern = @"(/)$";
            replacement = "";
            niceRow = Regex.Replace(niceRow, pattern, replacement);

            // Double up single quotes.
            pattern = @"([^']+)'([^']+)";
            replacement = "$1''$2";
            niceRow = Regex.Replace(niceRow, pattern, replacement);

            // Keep replacing until string length doesn't change.
            while (niceRow.Length > stringLength)
            {
                niceRow = Regex.Replace(niceRow, pattern, replacement);
                stringLength = niceRow.Length;
            }

            // Double up double quotes
            pattern = @"""([^""])";
            replacement = @"""""$1";
            stringLength = niceRow.Length;
            niceRow = Regex.Replace(niceRow, pattern, replacement);

            // Keep replacing until string length doesn't change.
            while (niceRow.Length > stringLength)
            {
                niceRow = Regex.Replace(niceRow, pattern, replacement);
                stringLength = niceRow.Length;
            }

            return niceRow;
        }

        /// <summary>
        /// Clears the "Invalid" highlighting & MouseOver eventhandler from a textbox.
        /// </summary>
        /// <param name="sheet">ActiveWorksheet.</param>

        internal static void ClearRegexInvalid(TextBox textBox)
        {
            if (textBox == null)
                return;

            // Clear any previous highlighting.
            textBox.BackColor = Color.White;

            // Remove the MouseHover EventHandler.
            DetachEvents(textBox);
        }

        /// <summary>
        /// Convert Excel-formatted date to SQL style.
        /// </summary>
        /// <param name="cellContents">String contents of a particular cell.</param>
        /// <returns>string</returns>
        internal static string ConvertExcelDate(string cellContents)
        {
            string convertedContents = null;

            if (!string.IsNullOrEmpty(cellContents))
            {
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
            }

            return convertedContents;
        }

        /// <summary>
        /// Removes event handlers from a text box.
        /// </summary>
        /// <param name="textBox">Handle to TextBox object</param>        
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

        /// <summary>
        /// Finds last column containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Columns, including cells that contains formulas that result in blank values
            return sheet.UsedRange.Columns.Count;
        }

        /// <summary>
        /// Finds last row containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastRow(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Row, including cells that contains formulas that result in blank values
            
            return sheet.UsedRange.Rows.Count;
        }

        /// <summary>
        /// Finds the ToolTip linked to a TextBox object.
        /// </summary>
        /// <param name="textBox">Handle to TextBox object</param>
        /// <returns>ToolTip</returns>
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

        /// <summary>
        /// Turn the scope-of-work filename into a .sql filename.
        /// </summary>
        /// <param name="filename">Scope of work filename</param>
        /// <param name="filenameAddOn">String we want to append to filename</param>
        /// <param name="filetype">Desired filetype (".sql" by default)</param>
        /// <param name="shortVersion">Bool--just filename.type? (false by default)</param>
        /// <returns>string</returns>
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

        /// <summary>
        /// Pulls all the column names from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>List<string></returns>
        internal static List<string> GetColumnNames(Worksheet sheet)
        {
            List<string> names = new List<string>();
            Range range = (Range)sheet.Cells[1, 1];

            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                try
                {
                    names.Add(range.Value.ToString());
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    break;
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return names;
        }

        /// <summary>
        /// Pulls all the column names from a worksheet.
        /// </summary>
        /// <param name="selectedColumns">List<Range></param>
        /// <returns>List<string></returns>
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

        /// <summary>
        /// Builds a dictionary of the column names from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>Dictionary mapping string -> Range</returns>
        internal static Dictionary<string, Range> GetColumnNamesDictionary(Worksheet sheet)
        {
            Dictionary<string, Range> columns = new Dictionary<string, Range>();
            Range range = (Range) sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                try
                {
                    columns.Add(range.Value.ToString(), range);
                }
                // If there's nothing in this header, then skip it.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                // If there's already a column by this name, skip this one.
                catch (System.ArgumentException) { }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return columns;
        }

        /// <summary>
        /// Has the user selected a column? And just one?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <param name="lastRow">Number of last row with data</param>
        /// <returns>Range</returns>
        internal static Range GetSelectedCol(Microsoft.Office.Interop.Excel.Application application, int lastRow)
        {
            Range selectedColumn = null;
            Range rng = (Range) application.Selection;

            // Whole column? Just one? And containing data?
            if (rng.Count > 1000 && 
                rng.Columns.Count == 1 && 
                Utilities.HasData(rng.Columns[1], lastRow))
            {
                // We want the TOP of the column.
                Worksheet sheet = application.Selection.Worksheet;
                int columnNumber = rng.Columns[1].Column;
                selectedColumn = (Range) sheet.Cells[1, columnNumber];
            }

            return selectedColumn;
        }

        /// <summary>
        /// Which columns has the user selected to export to SQL?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <param name="lastRow">Number of last row with data</param>
        /// <returns>List<Range></returns>
        internal static List<Range> GetSelectedCols(Microsoft.Office.Interop.Excel.Application application, int lastRow)
        {
            Range rng = (Range) application.Selection;
            List<Range> selectedColumns = new List<Range>();
            Worksheet sheet = application.Selection.Worksheet;

            foreach (Range col in rng.Columns)
            {
                // Don't add BLANK columns.
                if (Utilities.HasData(col, lastRow))
                {
                    // Want the TOP of the column.
                    int columnNumber = col.Column;
                    selectedColumns.Add((Range)sheet.Cells[1, columnNumber]);
                }
            }

            return selectedColumns;
        }

        /// <summary>
        /// Current time in yyyyMMddHHmmss format
        /// </summary>
        /// <returns>string</returns>
        // https://stackoverflow.com/q/21219797/18749636
        internal static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        internal static Dictionary<string, Worksheet> GetWorksheets()
        {
            Workbook workbook = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();
            
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                dict.Add(worksheet.Name, worksheet);
            }

            return dict;
        }

        /// <summary>
        /// Tests to see if RegEx pattern has any capture groups.
        /// </summary>
        /// <param name="regexText">Regular Expression</param>
        /// <returns>bool</returns>
        internal static bool HasCaptureGroups(string regexText)
        {
            bool hasCaptureGroups = false;

            // Empty strings are not errors.
            if (!string.IsNullOrEmpty(regexText))
            {
                try
                {
                    Regex regex = new Regex(regexText);

                    // https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.getgroupnumbers?view=net-8.0
                    int[] groupNumbers = regex.GetGroupNumbers();
                    hasCaptureGroups = groupNumbers.Count() > 1;
                }
                catch (ArgumentException)
                {
                }
            }

            return hasCaptureGroups;
        }

        /// <summary>
        /// Does this range have data?
        /// </summary>
        /// <param name="rng">Range to search</param>
        /// <param name="lastRow">Number of last row with data</param>
        /// <returns>bool</returns>
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
                    cell_contents = Convert.ToString(thisCell.Value2);
                    hasData = true;
                    break;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            }

            return hasData;
        }

        /// <summary>
        /// Inserts new column next to the provided Range.
        /// </summary>
        /// <param name="range">Range of existing column</param>
        /// <param name="newColumnName">Name of column to be created</param>
        /// <param name="side">Create it to the left or right of Range?</param>
        /// <returns>Range</returns>
        internal static Range InsertNewColumn(Range range, string newColumnName, InsertSide side = InsertSide.Right)
        {
            int columnNumber = range.Column;
            Worksheet sheet = range.Worksheet;
            Range newRange;

            if (side == InsertSide.Left)
            {
                sheet.Columns[columnNumber].EntireColumn.Insert();
                newRange = range.Offset[0, -1];
            }
            else
            {
                sheet.Columns[columnNumber + 1].EntireColumn.Insert();
                newRange = range.Offset[0, 1];
            }

            newRange.Value2 = newColumnName;
            return newRange;
        }

        /// <summary>
        /// Tests to see if string is a valid RegEx
        /// </summary>
        /// <param name="regexText">Regular Expression</param>
        /// <returns>RuleValidationResult object</returns>
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

        /// <summary>
        /// Highlight textbox to show its RegEx is not valid.
        /// </summary>
        /// <param name="textBox">TextBox object.</param>
        /// <param name="message">String used to fill the ToolTip.</param>
        
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

        // How many non-empty strings are present in the list?
        internal static int NumElementsPresent(List<string> values)
        {
            int numNonEmpties = 0;

            foreach (string value in values)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    numNonEmpties++;
                }
            }

            return numNonEmpties;
        }

        /// <summary>
        /// Open the output StreamWriter object, understanding that we might
        /// have to substitute a shorter version of the output filename
        /// if the default filename is too long.
        /// </summary>
        /// <param name="inputFilename">File to read</param>
        /// <param name="filenameAddOn">String we want to append to filename</param>
        /// <param name="filetype">Desired filetype (".sql" by default)</param>
        /// <returns>Tuple of StreamWriter object, string</returns>
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

        /// <summary>
        /// Saves the workbook as revised using a new name.
        /// </summary>
        /// <param name="workbook">Active workbook</param>
        /// <param name="newFilename">Desired new name for file</param>
        /// <param name="justTheFilename">Stub of filename in case we need to synthesize filename with timestamp</param>

        internal static void SaveRevised(Workbook workbook, string newFilename, string justTheFilename)
        {
            try
            {
                workbook.SaveCopyAs(newFilename);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                newFilename = System.IO.Path.Combine(
                    justTheFilename + "_" + Utilities.GetTimestamp()
                );
                workbook.SaveCopyAs(newFilename);
            }

            MessageBox.Show("Saved in '" + newFilename + "'.");
        }

        // Get the Range defined by these row, column Range objects.
        internal static Range ThisRowThisColumn(Range rowRange, Range columnRange)
        {
            Worksheet sourceSheet = rowRange.Worksheet as Worksheet;
            int rowNumber = rowRange.Row;

            int columnNumber = columnRange.Column;
            Range dataRange = (Range)sourceSheet.Cells[rowNumber, columnNumber];
            
            return dataRange;
        }

        /// <summary>
        /// Find the top cell in the named column.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <param name="columnName">Name of desired column.</param>
        /// <returns>Range</returns>
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

        /// <summary>
        /// Creates MessageBox letting user know we didn't find the named column.
        /// </summary>
        /// <param name="columnName">Name of desired column.</param>        
        internal static void WarnColumnNotFound(string columnName)
        {
            string message = "Column '" + columnName + "' not found.";
            string title = "Not Found";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);
        }
    }
}
