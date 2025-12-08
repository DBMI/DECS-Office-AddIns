using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;
using System.Windows.Forms;
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
    internal enum TypeOfMatch
    {
        DesiredNameStartsWithMatch,
        Exact,
        LastNameOnly,
        Levenshtein,
        MatchStartsWithDesiredName,
        UserSelected
    }
    internal class NameMatch
    {
        private string bestMatch;
        private TypeOfMatch matchType;
        private double? relativeDistance;

        internal NameMatch(string bestMatch, TypeOfMatch matchType, double? relativeDistance = null)
        {
            this.bestMatch = bestMatch;
            this.matchType = matchType;
            this.relativeDistance = relativeDistance;
        }

        internal string BestMatch()
        {
            return bestMatch;
        }
        internal bool IsMatch()
        {
            return !string.IsNullOrEmpty(bestMatch);
        }
        internal string MatchType()
        {
            string thisMatchType = matchType.ToString();

            if (matchType == TypeOfMatch.Levenshtein && relativeDistance.HasValue)
            {
                thisMatchType += ": " + relativeDistance.Value.ToString("F2");
            }

            return thisMatchType;
        }
        internal double? RelativeDistance()
        {
            return relativeDistance;
        }
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

            foreach (string columnName in columnNames)
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

        internal static void ClearWorksheet(Worksheet sheet)
        {
            Excel.Range usedRange = sheet.UsedRange;
            usedRange.Interior.Color = Constants.xlNone;
        }

        /// <summary>
        /// Finds all worksheets.
        /// </summary>
        /// <param name="workbook">Active Workbook.</param>

        internal static List<Worksheet> CollectAllWorksheets(Workbook workbook)
        {
            List<Worksheet> sheets = new List<Worksheet>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheets.Add(sheet);
            }

            return sheets;
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
        /// Given a list of column names ("Coverage Start Date", "Address Start Date"),
        /// finds the string they have in common ("Start Date").
        /// </summary>
        /// <param name="columnNames">List of strings</param>
        /// <returns>string</returns>
        internal static string CommonElements(List<string> columnNames)
        {
            string result = string.Empty;

            if (columnNames.Count == 0)
            {
                return result;
            }

            string longestSubstring = columnNames[0];

            for (int i = 1; i < columnNames.Count; i++)
            {
                longestSubstring = FindLongestCommonSubstring(longestSubstring, columnNames[i]);
            }

            return longestSubstring;
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
                catch (FormatException)
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
                catch (FormatException)
                {
                    // Try converting directly to DateTime.
                    if (DateTime.TryParse(cellContents, out DateTime result))
                    {
                        convertedContents = result;
                    }
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
        /// <param name="newName">string</param>

        public static Worksheet CreateNewNamedSheet(string newName)
        {
            int MAX_LENGTH = 31;
            Workbook workbook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;

            // Create new sheet at the end.
            Worksheet newSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);

            // There's a 31-character limit.
            string cleanName = newName;

            if (newName.Length > MAX_LENGTH)
                cleanName = newName.Substring(0, MAX_LENGTH);

            newSheet.Name = cleanName;
            return newSheet;
        }

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
                cellContents = column.Offset[rowOffset, 0].Value2.ToString();

                if (!result.Contains(cellContents))
                {
                    result.Add(cellContents);
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

        internal static List<string> ExtractColumnUnique(Range column)
        {
            List<string> names = new List<string>();
            int rowOffset = 1;

            while (true)
            {
                string cell_contents;

                try
                {
                    cell_contents = Convert.ToString(column.Offset[rowOffset, 0].Value2);

                    if (cell_contents.Length > 0 && !names.Contains(cell_contents))
                    {
                        names.Add(cell_contents);
                    }

                    rowOffset++;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // keep going
                }
                catch (NullReferenceException)
                {
                    break;
                }
            }

            names.Sort();
            return names;
        }

        internal static List<string> ExtractColumnUnique(Worksheet sheet, string colName)
        {
            List<string> names = new List<string>();
            int lastRow = Utilities.FindLastRow(sheet);
            Dictionary<string, Range> headers = Utilities.GetColumnRangeDictionary(sheet);

            if (headers.ContainsKey(colName))
            {
                Range thisCol = headers[colName];

                for (int rowOffset = 1; rowOffset < lastRow; rowOffset++)
                {
                    string cell_contents;

                    try
                    {
                        // Make them all upper case.
                        cell_contents = Convert.ToString(thisCol.Offset[rowOffset, 0].Value2).ToUpper();

                        if (cell_contents.Length > 0 && !names.Contains(cell_contents))
                        {
                            names.Add(cell_contents);
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                }
            }

            names.Sort();
            return names;
        }

        /// <summary>
        /// Finds the column name corresponding to a Range in a <name, Range> dictionary.
        /// </summary>
        /// <param name="columnNamesDict">Dictionary linking column names to ranges</param>
        /// <param name="columnRange">Range for which we want the column name</param>
        /// <returns>Range</returns>        

        internal static string FindColumnName(Dictionary<string, Range> columnNamesDict, Range columnRange)
        {
            int desiredColumnNumber = columnRange.Column;

            foreach (KeyValuePair<string, Range> entry in columnNamesDict)
            {
                int thisColumnNumber = entry.Value.Column;

                if (thisColumnNumber == desiredColumnNumber)
                {
                    return entry.Key;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Finds last column containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Detect Last used Columns, including cells that contain formulas that result in blank values
            return sheet.Cells.Find(
                                    "*",
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value,
                                    Excel.XlSearchOrder.xlByColumns,
                                    Excel.XlSearchDirection.xlPrevious,
                                    false,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value).Column;
        }

        /// <summary>
        /// Finds last row containing anything.
        /// </summary>
        /// <param name="rng">Range</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastRow(Range rng)
        {
            Worksheet sheet = rng.Worksheet;
            return sheet.Cells.Find(
                                    "*",
                                    System.Reflection.Missing.Value,
                                    Excel.XlFindLookIn.xlValues,
                                    Excel.XlLookAt.xlWhole,
                                    Excel.XlSearchOrder.xlByRows,
                                    Excel.XlSearchDirection.xlPrevious,
                                    false,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value).Row;
        }

        internal static int FindLastRow(Worksheet sheet)
        {
            return sheet.Cells.Find(
                                    "*",
                                    System.Reflection.Missing.Value,
                                    Excel.XlFindLookIn.xlValues,
                                    Excel.XlLookAt.xlWhole,
                                    Excel.XlSearchOrder.xlByRows,
                                    Excel.XlSearchDirection.xlPrevious,
                                    false,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value).Row;
        }

        /// <summary>
        /// Finds last worksheet.
        /// </summary>
        /// <param name="workbook">Active Workbook.</param>

        internal static Worksheet FindLastWorksheet(Workbook workbook)
        {
            List<Worksheet> sheets = new List<Worksheet>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheets.Add(sheet);
            }

            return sheets.LastOrDefault();
        }

        internal static NameMatch FindClosestMatch(List<string> names, string desiredName, double maxDistanceAllowed = 1.0)
        {
            double lowestScore = 1000000;
            string bestMatch = string.Empty;

            Fastenshtein.Levenshtein lev = new Fastenshtein.Levenshtein(desiredName);

            foreach (string thisName in names)
            {
                double wordLength = (double)Math.Min(desiredName.Length, thisName.Length);
                int levenshteinDistance = lev.DistanceFrom(thisName);
                double relativeDistance = levenshteinDistance / wordLength;

                if (thisName.StartsWith(desiredName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return new NameMatch(bestMatch: thisName, matchType: TypeOfMatch.MatchStartsWithDesiredName);
                }
                else if (desiredName.StartsWith(thisName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return new NameMatch(bestMatch: thisName, matchType: TypeOfMatch.DesiredNameStartsWithMatch);
                }
                else if (relativeDistance < lowestScore && relativeDistance < maxDistanceAllowed)
                {
                    bestMatch = thisName;
                    lowestScore = relativeDistance;
                }
            }

            return new NameMatch(bestMatch: bestMatch, matchType: TypeOfMatch.Levenshtein, relativeDistance: lowestScore);
        }

        internal static Rows FindNamesExact(Worksheet sheet, string namesColumnName, string desiredName)
        {
            desiredName = desiredName.ToUpper();
            Rows rows = new Rows();
            int lastRow = FindLastRow(sheet);
            Dictionary<string, Range> columnsDict = GetColumnRangeDictionary(sheet);

            if (columnsDict.ContainsKey(namesColumnName))
            {
                Range range = columnsDict[namesColumnName];
                string cellContents = string.Empty;

                for (int rowOffset = 0; rowOffset < lastRow; rowOffset++)
                {
                    cellContents = range.Offset[rowOffset, 0].Value2.ToString().ToUpper();

                    if (rows.HasStart())
                    {
                        // Then I'm looking for the next NON-match.
                        if (cellContents != desiredName)
                        {
                            rows.SetEnd(rowOffset);

                            // We're done.
                            break;
                        }
                    }
                    else
                    {
                        // Then I'm looking for the FIRST match.
                        if (cellContents == desiredName)
                        {
                            rows.SetStart(rowOffset + 1);
                        }
                    }
                }
            }

            // Did we run off the end of the sheet without finding a non-match?
            if (rows.HasStart() && !rows.HasEnd())
            {
                rows.SetEnd(lastRow);
            }

            return rows;
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

            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_offset = 0; col_offset < lastUsedCol; col_offset++)
            {
                string thisColumnName = string.Empty;

                try
                {
                    thisColumnName = Convert.ToString(range.Offset[0, col_offset].Value);
                    string test = Convert.ToString(range.Offset[0, col_offset].Value2);
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
                        columns.Add(thisColumnName, range.Offset[0, col_offset]);
                    }
                    // If there's already a column by this name, skip this one.
                    catch (System.ArgumentException) { }
                }
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

        internal static int GetIndexOfFirstWordAfterThis(List<string> words, string desiredWord)
        {
            int index = 0;

            foreach (string word in words)
            {
                if (string.Compare(word, desiredWord) < 0)
                {
                    index++;
                }
                else
                {
                    break;
                }
            }

            return index;
        }

        internal static List<string> GetListBoxContents(System.Windows.Forms.ListBox listBox)
        {
            List<string> contents = new List<string>();

            foreach (object item in listBox.Items)
            {
                // 'item' represents each item in the ListBox
                // You can cast it to the appropriate type if you know it (e.g., string)
                // For example: string listItemText = item.ToString();

                // You can then process each item as needed, for example, display it in a MessageBox
                contents.Add(item.ToString());
            }

            return contents;
        }

        /// <summary>
        /// Has the user selected a column? And just one?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <returns>Range</returns>
        internal static Range GetSelectedCol(Excel.Application application)
        {
            Range rng = (Range)application.Selection;
            Worksheet sheet = application.Selection.Worksheet;
            Range selectedColumn = null;

            // Whole column? Just one? And containing data?
            if (Utilities.HasData(rng.Columns[1]))
            {
                // We want the TOP of the column.
                int columnNumber = rng.Columns[1].Column;
                selectedColumn = (Range)sheet.Cells[1, columnNumber];
            }

            return selectedColumn;
        }

        /// <summary>
        /// Which columns has the user selected?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <param name="lastRow">Number of last row with data</param>
        /// <returns>List<Range></returns>
        internal static List<Range> GetSelectedCols(Microsoft.Office.Interop.Excel.Application application)
        {
            Range rng = (Range)application.Selection;
            List<Range> selectedColumns = new List<Range>();
            Worksheet sheet = application.Selection.Worksheet;

            foreach (Range col in rng.Columns)
            {
                // Don't add BLANK columns.
                if (Utilities.HasData(col))
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

        /// <summary>
        /// Finds all the worksheets.
        /// </summary>
        /// <returns>Dictionary<string, Worksheet></returns>
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
        /// <returns>bool</returns>
        internal static bool HasData(Range rng)
        {
            bool hasData = false;
            Range thisCell;
            int rowNumber = 0;

            while (true)
            {
                rowNumber++;
                thisCell = rng.Cells[rowNumber];
                string cell_contents;
                int numConsecutiveFailures = 0;

                try
                {
                    cell_contents = Convert.ToString(thisCell.Value2);
                    hasData = true;
                    break;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    numConsecutiveFailures++;

                    if (numConsecutiveFailures >= 3)
                    {
                        break;
                    }
                }
            }

            return hasData;
        }

        /// <summary>
        /// Given a Range to a column with a list of names ("Alice Apple", "Alice Apple", "Bob Baker"),
        /// finds the distinct elements ("Alice Apple", "Bob Baker") and the Block of rows where each appears.
        /// </summary>
        /// <param name="column">Range</param>
        /// <returns>Dictionary of Block objects</returns>
        internal static Dictionary<string, Block> IdentifyBlocks(Range column)
        {
            Dictionary<string, Block> dict = new Dictionary<string, Block>();
            string cellContents = string.Empty;
            Block thisBlock = null;
            int startingOffset = 1;
            int endingOffset = 1;
            string thisBlockName = null;
            int rowOffset = 0;

            while (true)
            {
                rowOffset++;
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
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Add the ending block.
                    thisBlock = new Block(startingOffset, endingOffset);

                    // Should we add this to the dictionary?
                    if (!dict.ContainsKey(thisBlockName))
                    {
                        dict[thisBlockName] = thisBlock;
                    }

                    break;
                }
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
                cellContents = column.Offset[rowOffset, 0].Value2.ToString();
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
        /// Finds longest common substring between two strings.
        /// </summary>
        /// <param name="s1">string</param>
        /// <param name="s2">string</param>
        /// <returns>string</returns>
        public static string FindLongestCommonSubstring(string s1, string s2)
        {
            string longestCommon = string.Empty;

            for (int i = 0; i < s1.Length; i++)
            {
                for (int j = i; j < s1.Length; j++)
                {
                    string currentSubstring = s1.Substring(i, j - i + 1);

                    if (s2.Contains(currentSubstring) && currentSubstring.Length > longestCommon.Length)
                    {
                        longestCommon = currentSubstring;
                    }
                }
            }
            return longestCommon;
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

        /// <summary>
        /// See if names in Last, First format at least match in last names.
        /// </summary>
        /// <param name="names">List<string>/string>.</param>
        /// <param name="desiredName">string</param>
        /// <returns>List<string></returns>
        internal static List<string> MightMatch(List<string> names, string desiredName)
        {
            List<string> possibleMatches = new List<string>();

            string[] pieces = desiredName.Split(',');
            string desiredLastName = pieces[0];

            foreach (string thisName in names)
            {
                string[] thesePieces = thisName.Split(',');

                if (string.Compare(desiredLastName, thesePieces[0], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    possibleMatches.Add(thisName);
                }
            }

            return possibleMatches;
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

        /// <summary>Open another Excel worksheet</summary>
        /// <param name="filename">Excel file to open</param>
        /// <returns>Workbook object</returns>

        internal static Workbook OpenExternalFile(string filename)
        {
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;

            if (xlApp != null)
            {
                try
                {
                    xlWorkbook = xlApp.Workbooks.Open(filename);
                }
                catch (Exception)
                {
                    xlApp.Quit();
                }
            }

            return xlWorkbook;
        }

        /// <summary>
        /// Open the output StreamWriter object, understanding that we might
        /// have to substitute a shorter version of the output filename
        /// if the default filename is too long.
        /// </summary>
        /// <param name="inputFilename">File to read</param>
        /// <param name="filenameAddOn">String we want to append to filename</param>
        /// <param name="filetype">Desired filetype (".sql" by default)</param>
        /// <param name="index">Number in a series (1 by default)</param>
        /// <returns>Tuple of StreamWriter object, string</returns>
        internal static (StreamWriter writer, string openedFilename) OpenOutput(
            string inputFilename,
            string filenameAddon = "",
            string filetype = ".sql",
            int index = 1
        )
        {
            // If we're creating a .sql file,
            // make the filename import-friendly by replacing spaces with underscores.
            bool replaceSpaces = filetype == ".sql";

            // Call the first file "file.sql" and the next one "file_2.sql".
            if (index > 1)
            {
                filenameAddon += "_" + index.ToString();
            }

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

        /// <summary>Fills a ListBox object with a list.</summary>
        /// <param name="listBox">ListBox object</param>
        /// <param name="contents">List<string></param>
        internal static void PopulateListBox(System.Windows.Forms.ListBox listBox, List<string> contents, bool enableWhenPopulated = false)
        {
            listBox.Items.Clear();

            foreach (string item in contents)
            {
                listBox.Items.Add(item);
            }

            if (enableWhenPopulated)
            {
                listBox.Enabled = true;
            }
        }

        /// <summary>
        /// Looks for the given string in ANY row of column.
        /// </summary>
        /// <param name="column">Range</param>
        /// <param name="target">Range</param>
        /// <returns>bool</returns>
        internal static bool PresentInColumn(Range column, string target)
        {
            bool foundIt = false;
            Worksheet sheet = column.Worksheet;
            int lastUsedRow = sheet.UsedRange.Rows.Count;

            Range searchedCell = (Range)sheet.Cells[1, column.Column];

            for (int rowOffset = 0; rowOffset < lastUsedRow; rowOffset++)
            {
                try
                {
                    string contents = searchedCell.Offset[rowOffset, 0].Value.ToString();

                    if (contents.Contains(target))
                    {
                        foundIt = true;
                        break;
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            }

            return foundIt;
        }

        /// <summary>
        /// Looks for the given string in the first row of provided column.
        /// </summary>
        /// <param name="column">Range</param>
        /// <param name="target">Range</param>
        /// <returns>bool</returns>
        internal static bool PresentInHeader(Range column, string target)
        {
            bool foundIt = false;
            Worksheet sheet = column.Worksheet;
            int lastUsedRow = sheet.UsedRange.Rows.Count;

            // Don't search to the bottom--just the top few rows.
            // (But don't accidentally search past the bottom of the data, either.)
            int rowsToSearch = Math.Min(3, lastUsedRow);

            Range searchedCell = (Range)sheet.Cells[1, column.Column];

            for (int rowOffset = 0; rowOffset < rowsToSearch; rowOffset++)
            {
                try
                {
                    string contents = searchedCell.Offset[rowOffset, 0].Value.ToString();

                    if (contents.Contains(target))
                    {
                        foundIt = true;
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            }

            return foundIt;
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

        /// <summary>
        /// Scrolls the worksheet to desired row.
        /// </summary>
        /// <param name="worksheet">Active worksheet</param>
        /// <param name="rowNumber">int</param>
        internal static void ScrollToRow(Worksheet worksheet, int rowNumber)
        {
            Excel.Application application = Globals.ThisAddIn.Application;

            try
            {
                // Don't try to scroll above row 1.
                rowNumber = Math.Max(rowNumber, 1);

                // Don't try to scroll past the end of the sheet.
                rowNumber = Math.Min(rowNumber, worksheet.Rows.Count);

                application.ActiveWindow.ScrollRow = rowNumber;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error scrolling to row: {ex.Message}");
            }
        }

        internal static void SortPageByColumn(Worksheet sheet, Range selectedColumnRng)
        {
            // Clear existing sort fields (optional but good practice)
            sheet.Sort.SortFields.Clear();

            // Add a sort field.
            sheet.Sort.SortFields.Add(
                Key: selectedColumnRng, // first data cell in the sort key column
                SortOn: Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues,
                Order: Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                DataOption: Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal
            );

            // Set the range to be sorted.
            int lastRow = FindLastRow(sheet);
            Range wholeColumn = sheet.Range[selectedColumnRng, selectedColumnRng.Offset[lastRow - 1, 0]];
            sheet.Sort.SetRange(wholeColumn); // should include the header if 'Header' is xlYes

            // Apply the sort
            sheet.Sort.Header = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes; // Set to xlNo if no header
            sheet.Sort.MatchCase = false;
            sheet.Sort.Orientation = Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns; // Or xlSortRows
            sheet.Sort.Apply();
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

        /// <summary>
        /// Tests to see what fraction of the words in "name" are present in another name.
        /// Ex. Tests "Dr. Able Wise" vs. "Wise, Able MD" ==> 0.667
        /// </summary>
        /// <param name="name">string.</param>
        /// <param name="otherName">string.</param>
        /// <returns>double</returns>
        internal static double WordsPresent(string name, string otherName)
        {
            string[] wordsInNewName = name.Replace(",", "").Trim().Split();
            string[] wordsInOtherName = otherName.Replace(",", "").Trim().Split();
            double numWordsToTest = (double)wordsInNewName.Length;
            double wordsPresent = 0.0;

            foreach (string word in wordsInNewName)
            {
                if (wordsInOtherName.Contains(word))
                {
                    wordsPresent++;
                }
            }

            return wordsPresent / numWordsToTest;
        }
    }
}
