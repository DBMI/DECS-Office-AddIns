using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
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
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

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
            
            pattern = @"([^""])""$";
            replacement = "$1";
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
        /// <param name="textBox">TextBox object</param>

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
        /// Combine multiple columns across one row.
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="rowNumber">int Row being processed</param>
        /// <param name="columns">List<Range> Columns being combined</param>
        /// <returns>string</returns>
        internal static string CombineColumns(Worksheet sheet, int rowNumber, List<Range> columns)
        {
            string textCombined = string.Empty;
            int columnNumber;
            Range source;

            foreach (Range column in columns)
            {
                columnNumber = column.Column;
                source = (Range)sheet.Cells[rowNumber, columnNumber];

                try
                {
                    textCombined = textCombined + source.Value.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                { 
                    // Cell is probably empty.
                }
            }

            return textCombined;
        }

        /// <summary>
        /// Convert Excel-formatted date to SQL style.
        /// </summary>
        /// <param name="cellContents">String contents of a particular cell.</param>
        /// <returns>string</returns>
        internal static string ConvertExcelDateToString(string cellContents)
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
        /// Convert Excel-formatted date to SQL style.
        /// </summary>
        /// <param name="cellContents">String contents of a particular cell.</param>
        /// <returns>DateTime</returns>
        internal static DateTime? ConvertExcelDate(string cellContents)
        {
            DateTime? convertedContents = null;

            if (!string.IsNullOrEmpty(cellContents))
            {
                try
                {
                    double d = double.Parse(cellContents);
                    convertedContents = DateTime.FromOADate(d);
                }
                catch (System.FormatException)
                {
                    // Probably trying to convert the name "Date" to a Double in order to create DateTime object.
                }
            }

            return convertedContents;
        }

        /// <summary>
        /// Copy block of rows from source Worksheet to target Worksheet.
        /// </summary>
        /// <param name="sourceSheet">Worksheet</param>
        /// <param name="sourceBlock">Block</param>
        /// <param name="targetSheet">Worksheet</param>
        /// <param name="targetBlock">Block</param>

        public static void CopyBlock(Worksheet sourceSheet, Block sourceBlock, Worksheet targetSheet, Block targetBlock)
        {
            Range sourceRange = sourceSheet.Rows[sourceBlock.rowList()];
            Range targetRange = targetSheet.Rows[targetBlock.rowList()];
            sourceRange.Copy(targetRange);
        }

        /// <summary>
        /// Copy row from source Worksheet to target Worksheet.
        /// </summary>
        /// <param name="sourceSheet">Worksheet</param>
        /// <param name="sourceRowOffset">int</param>
        /// <param name="targetSheet">Worksheet</param>
        /// <param name="targetRowOffset">int</param>

        public static void CopyRow(Worksheet sourceSheet, int sourceRowOffset, Worksheet targetSheet, int targetRowOffset)
        {
            // Convert from offset to row number.
            int sourceRowNumber = sourceRowOffset + 1;
            int targetRowNumber = targetRowOffset + 1;

            Range sourceRange = sourceSheet.Rows[sourceRowNumber + ":" + sourceRowNumber];
            Range targetRange = targetSheet.Rows[targetRowNumber + ":" + targetRowNumber];
            sourceRange.Copy(targetRange);
        }

        /// <summary>
        /// Insert new Worksheet with given name.
        /// </summary>
        /// <param name="sourceSheet">Worksheet</param>
        /// <param name="newName">string</param>

        public static Worksheet CreateNewNamedSheet(Worksheet worksheet, string newName)
        {        
            int MAX_LENGTH = 31;
            Workbook workbook = worksheet.Parent;

            // Create new sheet at the end.
            Worksheet newSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);

            // There's a 31-character limit.
            string cleanName = newName;

            if (newName.Length > MAX_LENGTH)
                cleanName = newName.Substring(0, MAX_LENGTH);

            newSheet.Name = cleanName;
            return newSheet;
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
        /// Given a list of names ("Alice Apple", "Alice Apple", "Bob Baker"),
        /// finds the distinct elements ("Alice Apple", "Bob Baker").
        /// </summary>
        /// <param name="names">List of strings</param>
        /// <returns>List of strings</returns>
        internal static List<string> Distinct(List<string> names)
        {
            List<string> result = new List<string>();

            foreach (string name in names)
            {
                if (!result.Contains(name))
                {
                    result.Add(name);
                }
            }

            result.Sort();
            return result;
        }

        /// <summary>
        /// Given a Range to a column with a list of names ("Alice Apple", "Alice Apple", "Bob Baker"),
        /// finds the distinct elements ("Alice Apple", "Bob Baker").
        /// </summary>
        /// <param name="column">Range</param>
        /// <param name="lastRow">int</param>
        /// <returns>List of strings</returns>
        internal static List<string> Distinct(Range column, int lastRow)
        {
            List<string> result = new List<string>();
            string cellContents = string.Empty;

            for (int rowOffset = 1; rowOffset < (lastRow - 1); rowOffset++)
            {
                try
                {
                    cellContents = column.Offset[rowOffset, 0].Value2.ToString();

                    if (!result.Contains(cellContents))
                    {
                        result.Add(cellContents);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            result.Sort();
            return result;
        }

        /// <summary>
        /// Given a list of column names ("Coverage Start Date", "Address Start Date"),
        /// finds the strings that make them different ("Address", "Coverage").
        /// </summary>
        /// <param name="columnNames">List of strings</param>
        /// <param name="ignoredWords">List of strings to ignore, like "Start", "Date"</param>
        /// <returns>List of strings</returns>
        internal static List<string> DistinctElements(List<string> columnNames, List<string> ignoredWords)
        {
            List<string> result = new List<string>();

            // Break up all the column names into a list of words, containing duplicates.
            List<string> pieces = new List<string>();

            foreach (string columnName in columnNames)
            {
                pieces = columnName.Split().ToList();

                foreach (string piece in pieces)
                {
                    if (!ignoredWords.Contains(piece) && !result.Contains(piece))
                    {
                        result.Add(piece);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Finds last column containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Not sure why (at some point) I thought it was necessary to do this??
            // Unhide All Cells and clear formats
            //sheet.Columns.ClearFormats();
            //sheet.Rows.ClearFormats();

            // Detect Last used Columns, including cells that contain formulas that result in blank values
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
            //sheet.Columns.ClearFormats();
            //sheet.Rows.ClearFormats();

            // Detect Last used Row, including cells that contain formulas that result in blank values
            
            return sheet.UsedRange.Rows.Count;
        }

        internal static Worksheet FindLastWorksheet(Workbook workbook)
        {
            List<Worksheet> sheets = new List<Worksheet>();
            
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheets.Add(sheet);
            }

            return sheets.LastOrDefault();
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
        /// <param name="replaceSpaces">Should we replace spaces with underscores (true by default)</param>
        /// <param name="shortVersion">Bool--just filename.type? (false by default)</param>
        /// <returns>string</returns>
        internal static string FormOutputFilename(
            string filename,
            string filenameAddon = "",
            string filetype = ".sql",
            bool replaceSpaces = true,
            bool shortVersion = false
        )
        {
            string dir = Path.GetDirectoryName(filename);
            string justTheFilename = Path.GetFileNameWithoutExtension(filename) + filenameAddon;

            if (replaceSpaces)
            {
                // Make SQL filename import-friendly by replacing spaces with underscores.
                justTheFilename = justTheFilename.Replace(' ', '_');
            }

            string sqlFilename = Path.Combine(dir, justTheFilename + filetype);

            if (shortVersion)
            {
                sqlFilename = justTheFilename + filetype;
            }

            return sqlFilename;
        }

        /// <summary>
        /// Reads the column name from the first row of column Range.
        /// </summary>
        /// <param name="column">Range</param>
        /// <returns>string</returns>
        internal static string GetColumnName(Range column)
        {
            string name = column.Offset[0, 0].Value2.ToString();
            return name;
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
        /// Builds a dictionary linking column names to ranges from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet</param>
        /// <param name="namesDesired">List<string></string></param>
        /// <param name="caseSensitive">bool</param>
        /// <returns>Dictionary mapping string -> Range</returns>
        internal static Dictionary<string, Range> GetColumnRangeDictionary(Worksheet sheet, 
                                                                           List<string> namesDesired = null,
                                                                           bool caseSensitive = true)
        {
            Dictionary<string, Range> columns = null;

            if (caseSensitive)
            {
                columns = new Dictionary<string, Range>();
            }
            else
            {
                var comparer = StringComparer.OrdinalIgnoreCase;
                columns = new Dictionary<string, Range>(comparer);
            }

            Range range = (Range) sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                string thisColumnName = string.Empty;

                try
                {
                    thisColumnName = Convert.ToString(range.Value);
                }
                // If there's nothing in this header, move to next column.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) 
                {
                    continue;
                }

                // If column name is empty, can't do anything.
                if (string.IsNullOrEmpty(thisColumnName))
                {
                    continue;
                }

                // If namesDesired isn't specified, then we add every column name.
                // Otherwise, see if this column is one of the droids we're looking for.
                if (namesDesired is null || namesDesired.Count == 0 || namesDesired.Contains(thisColumnName))
                {
                    try
                    {
                        columns.Add(thisColumnName, range);
                    }
                    // If there's already a column by this name, skip this one.
                    catch (System.ArgumentException) { }
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return columns;
        }

        /// <summary>
        /// Builds a dictionary linking column names to ColumnType enum from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>Dictionary mapping string -> ColumnType</returns>
        internal static Dictionary<string, ColumnType> GetColumnTypeDictionary(Worksheet sheet)
        {
            Dictionary<string, ColumnType> columns = new Dictionary<string, ColumnType>();
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                try
                {
                    string columnName = range.Value.ToString();
                    ColumnType columnType = ColumnType.Text;

                    if (columnName.Contains("Date") || columnName.Contains("DTTM"))
                    {
                        columnType = ColumnType.Date;
                    }
                    
                    columns.Add(range.Value.ToString(), columnType);
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
        /// Given a Range to a column with a list of names ("Alice Apple", "Alice Apple", "Bob Baker"),
        /// finds the distinct elements ("Alice Apple", "Bob Baker") and the Block of rows where each appears.
        /// </summary>
        /// <param name="column">Range</param>
        /// <param name="lastRow">int</param>
        /// <returns>Dictionary of Block objects</returns>
        internal static Dictionary<string, Block> IdentifyBlocks(Range column, int lastRow)
        {
            Dictionary<string, Block> dict = new Dictionary<string, Block>();
            string cellContents = string.Empty;
            Block thisBlock = null;
            int startingOffset = 1;
            int endingOffset = 1;
            string thisBlockName = null;

            for (int rowOffset = 1; rowOffset < lastRow; rowOffset++)
            {
                try
                {
                    cellContents = column.Offset[rowOffset, 0].Value2.ToString();

                    if (thisBlockName is null || cellContents == thisBlockName)
                    {
                        // Still in same block, so keep a running count.
                        thisBlockName = cellContents;
                        endingOffset = rowOffset;
                    }
                    else
                    {
                        // The name just changed, which means:
                        //  --> the previous block ended.
                        thisBlock = new Block(startingOffset, endingOffset);

                        // Should we add this to the dictionary?
                        if (!dict.ContainsKey(thisBlockName))
                        {
                            dict[thisBlockName] = thisBlock;
                        }

                        //  --> & a new block started.
                        thisBlockName = cellContents;
                        startingOffset = rowOffset;
                        endingOffset = rowOffset;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            // Add the ending block.
            thisBlock = new Block(startingOffset, endingOffset);

            // Should we add this to the dictionary?
            if (!dict.ContainsKey(thisBlockName))
            {
                dict[thisBlockName] = thisBlock;
            }

            return dict;
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
        /// Tests to see if column can be converted to Excel date.
        /// </summary>
        /// <param name="cellContents">String contents of a particular cell.</param>
        /// <returns>DateTime</returns>
        internal static bool IsExcelDate(Range column, int lastRow)
        {
            bool isExcelDate = true;
            string cellContents = string.Empty;
            DateTime pastDate = DateTime.Parse("1950-01-01");
            DateTime futureDate = DateTime.Parse("2050-01-01");

            for (int rowOffset = 1; rowOffset < (lastRow - 1); rowOffset++)
            {
                try
                {
                    cellContents = column.Offset[rowOffset, 0].Value2.ToString();
                }
                catch (Exception ex)
                { 
                    MessageBox.Show(ex.Message);
                }

                DateTime? convertedDate = ConvertExcelDate(cellContents);

                if (convertedDate == null || convertedDate < pastDate || convertedDate > futureDate)
                {
                    isExcelDate = false;
                    break;
                }
            }

            return isExcelDate;
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
            // If we're creating a .sql file,
            // make the filename import-friendly by replacing spaces with underscores.
            bool replaceSpaces = filetype == ".sql";

            string outputFilename = Utilities.FormOutputFilename(
                filename: inputFilename,
                filetype: filetype,
                filenameAddon: filenameAddon,
                replaceSpaces: replaceSpaces,
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

        internal static void PopulateListBox(System.Windows.Forms.ListBox listBox, List<string> contents)
        {
            listBox.Items.Clear();

            foreach (string item in contents)
            {
                listBox.Items.Add(item);
            }
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
