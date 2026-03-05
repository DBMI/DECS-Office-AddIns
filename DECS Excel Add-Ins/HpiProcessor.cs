using MathNet.Numerics.Distributions;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DECS_Excel_Add_Ins
{    /**
     * @brief Main class for @c AddHPI tool.
     */

    internal class HpiProcessor
    {
        private Application application;
        private const int HALFWAY_DOWN_THE_SHEET = 12;

        internal HpiProcessor()
        {
            application = Globals.ThisAddIn.Application;
        }

        /// <summary>
        /// Finds either the user-selected column or (if none selected) column with name we expect.
        /// <summary>
        /// <param name="worksheet">Reference to the ActiveSheet.</param>
        /// <param name="desiredName">Name of column we're looking for.</param>
        /// <returns>Range</returns>
        private Range FindNamedColumn(Worksheet worksheet, string desiredName)
        {
            Regex desiredPattern = new Regex(desiredName.ToLower());
            Range selectedColumn = Utilities.GetSelectedCol(application);

            // If user didn't select a column, find it by name.
            if (selectedColumn == null)
            {
                Dictionary<string, Range> columns = Utilities.GetColumnRangeDictionary(worksheet);

                foreach (KeyValuePair<string, Range> column in columns)
                {
                    Match match = desiredPattern.Match(column.Key.ToLower());

                    if (match.Success)
                    {
                        selectedColumn = column.Value;
                        break;
                    }
                }
            }
            else
            {
                // What's the heading of this column say?
                string header = selectedColumn.Value2;
                Match match = desiredPattern.Match(header.ToLower());

                if (!match.Success)
                {
                    return null;
                }
            }

            return selectedColumn;
        }

        /// <summary>
        /// Scans the worksheet:
        /// - Finds the address column (or the zip column, if address not found),
        ///    either using the selected column or finding it by name.
        /// - Reads data file SVI_2020_US.csv & populates a dictionary mapping census tract # to SVI values.
        /// - Uses online geocode service to retrieve the census tract for each address.
        /// - Looks up the SVI values from the tract dictionary.
        /// <summary>
        /// <param name="worksheet">Reference to the ActiveSheet.</param>

        internal void Scan(Worksheet worksheet)
        {
            // 1) Find the census tract column.
            Range locationColumn = FindNamedColumn(worksheet, "Census FIPS");

            if (locationColumn == null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        locationColumn = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return;
                    }
                }
            }

            // 2) Populate the HPI dictionary from data file.
            application.StatusBar = "Building HPI dictionary.";
            HpiTable hpiTable = new HpiTable();

            if (hpiTable.ready)
            {
                // Build output columns.
                Range hpiPercentileColumn = Utilities.InsertNewColumn(range: locationColumn, newColumnName: "HPI percentile", side: InsertSide.Right);
                Range hpiScoreColumn = Utilities.InsertNewColumn(range: locationColumn, newColumnName: "HPI score", side: InsertSide.Right);

                int rowOffset = 1;
                int numConsecutiveFailures = 0;

                // 3) Convert each census tract FIPS number to HPI.
                while (true)
                {
                    try
                    {
                        string location = locationColumn.Offset[rowOffset, 0].Text;

                        if (string.IsNullOrEmpty(location))
                        {
                            numConsecutiveFailures++;
                        }
                        else
                        {
                            if (ulong.TryParse(location, out ulong fips))
                            {                            
                                // Don't display nonsense numbers (represented by -1).
                                double? rawScore = hpiTable.hpi(fips);

                                if (rawScore.HasValue)
                                {
                                    hpiScoreColumn.Offset[rowOffset, 0].Value2 = rawScore.Value;
                                }

                                double? percentile = hpiTable.hpi_percentile(fips);

                                if (percentile.HasValue)
                                {
                                    hpiPercentileColumn.Offset[rowOffset, 0].Value2 = percentile.Value;
                                }

                                // reset
                                numConsecutiveFailures = 0;
                            }
                            else
                            {
                                numConsecutiveFailures++;
                            }
                        }
                    }
                    catch
                    {
                        break;
                    }

                    // An occasional miss is ok, but three in a row & we've run outta data.
                    if (numConsecutiveFailures >= 3)
                    {
                        break;
                    }

                    rowOffset++;
                    Utilities.ScrollToRow(worksheet, locationColumn.Offset[rowOffset].Row - HALFWAY_DOWN_THE_SHEET);

                    if (rowOffset % 10 == 0)
                    {
                        application.StatusBar = "Processed " + rowOffset.ToString() + " locations.";
                    }
                }

                application.StatusBar = "Complete";
            }
        }
    }
}
