using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Some information needs to be encoded in SQL as a date, the rest can use varchar.
    */
    // https://stackoverflow.com/a/479417/18749636
    internal enum DataType
    {
        [Description("date")]
        Date,

        // How we need to describe data type in SQL import statement.
        [Description("varchar(18)")]
        Varchar
    }

    /**
     * @brief Handles importing an Excel column of MRN, ICD codes, etc. into SQL format.
     */
    internal class ListImporter
    {
        private Application application;
        private List<string> columnNames;
        private int lastRow;
        private List<string> sqlVariableNames;
        private IDictionary<string, DataType> supportedDataTypes;

        private const int MAX_LINES_PER_IMPORT = 1000;

        // If these cells are empty, we'll skip the rest of the columns.
        private readonly string[] INDEX_ROW_NAMES = { "MRN", "PAT_ID" };

        private const string MAIN_TABLE_CREATE =
            "DROP TABLE IF EXISTS #PATIENT_LIST;\r\nCREATE TABLE #PATIENT_LIST (";
        private const string MAIN_TABLE_USE =
            ":setvar path \"F:\\DECS\\<task folder name>\"\r\n:r $(path)\\";
        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";
        private const string QUOTE = "'";
        private const string SEGMENT_START_I = "INSERT INTO #DATA_LIST (";
        private const string SEGMENT_START_II = ")\r\nVALUES\r\n";

        /// <summary>
        /// Constructor
        /// </summary>
        
        internal ListImporter()
        {
            application = Globals.ThisAddIn.Application;

            // Initialize dictionary to translate column names like "Date of Procedure" to DataType.Date.
            supportedDataTypes = new Dictionary<string, DataType>
            {
                { "Date", DataType.Date }
            };
        }

        /// <summary>
        /// Turn the column name (in row 1) into a enum data type.
        /// </summary>
        /// <param name="col">Range of column to search</param>
        /// <returns>DataType</returns>
        private DataType DetermineDataType(Range col)
        {
            DataType dataType = DataType.Varchar;

            try
            {
                // What's in the top cell?
                string colName = col.Cells[1].Value2.ToString();

                dataType = NameToDataType(colName);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return dataType;
            }

            return dataType;
        }

        /// <summary>
        /// Build a list of the cell contents across this row.
        /// </summary>
        /// <param name="columns">List<Range> of columns to search</param>
        /// <param name="rowNum">int row number to search</param>
        /// <returns>List<string></returns>
        private List<string> ExtractRow(List<Range> columns, int rowNum)
        {
            List<string> rowContents = new List<string>();

            foreach (Range col in columns)
            {
                DataType dataType = DetermineDataType(col);

                Range thisCell = col.Cells[rowNum];
                string cellContents;

                try
                {
                    cellContents = thisCell.Value2.ToString();

                    // If the line is just the column name, skip this row.
                    if (columnNames.Any(cellContents.Contains))
                        break;

                    switch (dataType)
                    {
                        // Turn dates into something SQL will understand.
                        case DataType.Date:
                            cellContents = Utilities.ConvertExcelDate(cellContents);
                            break;

                        case DataType.Varchar:
                            cellContents = Utilities.CleanDataForSQL(cellContents);
                            break;

                        default:
                            break;
                    }

                    // Put quotes here because we DON'T want to wrap Null in quotes.
                    cellContents = QUOTE + cellContents + QUOTE;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // There's nothing in this cell.
                    // If it's an index column, skip the whole row.
                    if (IsIndexColumn(col))
                    {
                        break;
                    }

                    // Else leave a placeholder.
                    cellContents = "NULL";
                }

                rowContents.Add(cellContents);
            }

            return rowContents;
        }

        /// <summary>
        /// Based on the column name (in row 1) is this a special "index" column?
        /// </summary>
        /// <param name="col">Range of column to search</param>
        /// <returns>bool</returns>
        private bool IsIndexColumn(Range col)
        {
            bool isIndexColumn = false;

            try
            {
                // What's in the top cell?
                string colName = col.Cells[1].Value2.ToString();

                isIndexColumn = INDEX_ROW_NAMES.Any(colName.Contains);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            return isIndexColumn;
        }

        /// <summary>
        /// Figure out which DataType to use based on the column's name.
        /// </summary>
        /// <param name="colName">string column name</param>
        /// <returns>DataType</returns>
        private DataType NameToDataType(string colName)
        {
            DataType dataType = DataType.Varchar;

            foreach (KeyValuePair<string, DataType> entry in supportedDataTypes)
            {
                // Make case-insensitive match.
                if (colName.ToLower().Contains(entry.Key.ToLower()))
                {
                    dataType = entry.Value;
                    break;
                }
            }

            return dataType;
        }

        /// <summary>
        /// Initializes the SQL INSERT INTO statement.
        /// </summary>
        /// <param name="worksheet">Active worksheet</param>
        /// <returns>List<Range>, string</returns>
        private (List<Range> columns, string segmentStart) PrepSegmentStart(Worksheet worksheet)
        {
            string segmentStart = SEGMENT_START_I;

            // Any columns selected?
            List<Range> selectedColumns = Utilities.GetSelectedCols(application, lastRow);

            if (selectedColumns.Count == 0)
            {
                // Just take the first column & hope for the best!
                selectedColumns.Add(worksheet.Columns[1].EntireColumn);
            }

            // Build the list of column names.
            columnNames = Utilities.GetColumnNames(selectedColumns);

            sqlVariableNames = new List<string>();

            foreach (string column in columnNames)
            {
                sqlVariableNames.Add(Utilities.CleanColumnNamesForSQL(column));
            }

            // Turn "Date of consult" into "DATE_OF_CONSULT".
            segmentStart += string.Join(", ", sqlVariableNames);
            segmentStart += SEGMENT_START_II;
            return (selectedColumns, segmentStart);
        }

        private string PrepSegmentStart(List<string> externalColumnNames)
        {
            sqlVariableNames = new List<string>();

            foreach(string column in externalColumnNames)
            {
                sqlVariableNames.Add(Utilities.CleanColumnNamesForSQL(column));
            }

            string segmentStart = SEGMENT_START_I;

            // Turn "Date of consult" into "DATE_OF_CONSULT".
            segmentStart += string.Join(", ", sqlVariableNames);
            segmentStart += SEGMENT_START_II;
            return segmentStart;
        }

        /// <summary>
        /// Scans the worksheet & creates the SQL file that lists the patient data to be imported.
        /// </summary>
        /// <param name="worksheet">Active worksheet</param>

        internal void Scan(Worksheet worksheet)
        {
            // We'll use this in a lot of places, so let's just look it up once.
            lastRow = Utilities.FindLastRow(worksheet);

            if (lastRow == 1)
            {
                // Then perhaps the user wants to read/parse an external file.
                ScanExternalFile();
            }
            else
            {
                // Then the data are present on this sheet.
                ScanWorksheet(worksheet);
            }
        }

        private void ScanExternalFile()
        {
            // What's the source file?
            string externalFilename = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Because we don't specify an opening directory,
                // the dialog will open in the last directory used.
                openFileDialog.Filter = "csv files (*.csv)|*.csv";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified file.
                    externalFilename = openFileDialog.FileName;
                }
            }

            if (externalFilename != string.Empty)
            {

                (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                    inputFilename: externalFilename,
                    filenameAddon: "_list",
                    filetype: ".sql"
                );

                int lines_written_this_chunk = 0;

                using (TextFieldParser csvParser = new TextFieldParser(externalFilename))
                {
                    csvParser.CommentTokens = new string[] { "#" };
                    csvParser.SetDelimiters(new string[] { "\t" });
                    csvParser.HasFieldsEnclosedInQuotes = false;

                    // Read the row with the column names
                    List<string> columnHeaders = csvParser.ReadFields().ToList<string>();

                    // Use the column names to write the file header.
                    string segmentStart = PrepSegmentStart(columnHeaders);
                    writer.Write(PREAMBLE + segmentStart);

                    while (!csvParser.EndOfData)
                    {
                        // Read current line fields, pointer moves to the next line.
                        List<string> rowContents = csvParser.ReadFields().ToList<string>();
                        writer.Write("(" + QUOTE + string.Join(QUOTE + ", " + QUOTE, rowContents) + QUOTE + ")");
                        lines_written_this_chunk++;
                        string line_ending;

                        if (lines_written_this_chunk < MAX_LINES_PER_IMPORT)
                        {
                            line_ending = ",\r\n";
                        }
                        else
                        {
                            line_ending = ";\r\n\r\n" + segmentStart;
                            lines_written_this_chunk = 0;
                        }
                        writer.Write(line_ending);
                    }

                    writer.Write(";\r\n");
                }

                writer.Close();
                Process.Start(outputFilename);
            }
        }

        private void ScanWorksheet(Worksheet worksheet)
        {
            // Initialize the output .SQL file.
            Workbook workbook = worksheet.Parent;
            string workbookFilename = workbook.FullName;

            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                inputFilename: workbookFilename,
                filenameAddon: "_list",
                filetype: ".sql"
            );

            var selection = PrepSegmentStart(worksheet);
            writer.Write(PREAMBLE + selection.segmentStart);

            int lines_written_this_chunk = 0;
            application.StatusBar = "Processing...";

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
            {
                List<string> rowContents = ExtractRow(selection.columns, rowNumber);

                if (!rowContents.Any())
                {
                    continue;
                }

                writer.Write("(" + string.Join(", ", rowContents) + ")");
                lines_written_this_chunk++;
                string line_ending;

                if (rowNumber == lastRow)
                {
                    line_ending = ";\r\n";
                }
                else
                {
                    if (lines_written_this_chunk < MAX_LINES_PER_IMPORT)
                    {
                        line_ending = ",\r\n";
                    }
                    else
                    {
                        line_ending = ";\r\n\r\n" + selection.segmentStart;
                        lines_written_this_chunk = 0;
                    }
                }

                writer.Write(line_ending);

                if (rowNumber % 100 == 0)
                {
                    application.StatusBar = "Processed " + rowNumber.ToString() + "/" + lastRow.ToString() + " rows.";
                }
            }

            application.StatusBar = "Completed";
            writer.Close();
            Process.Start(outputFilename);
            WriteMainHeader(workbookFilename);
        }

        /// <summary>
        /// Writes the part of the main SQL script that creates a temp table from the patient list file.
        /// </summary>
        /// <param name="filename">Name of output file</param>
        
        private void WriteMainHeader(string filename)
        {
            // Build list of variables & types like "PAT_ID varchar, PROCEDURE_DATE date"
            List<string> variableNamesAndTypes = new List<string>();

            foreach (string varName in sqlVariableNames)
            {
                DataType dataType = NameToDataType(varName);
                variableNamesAndTypes.Add(varName + " " + dataType.GetDescription());
            }

            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                inputFilename: filename,
                filetype: ".sql"
            );

            writer.Write(PREAMBLE);
            writer.Write(MAIN_TABLE_CREATE + string.Join(", ", variableNamesAndTypes) + ")\r\n");
            string justTheFilenameAndExt = Path.GetFileName(outputFilename);
            writer.Write(MAIN_TABLE_USE + justTheFilenameAndExt + "\r\n");
            writer.Close();
            Process.Start(outputFilename);
        }
    }
}
