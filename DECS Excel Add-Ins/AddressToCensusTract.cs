using MathNet.Numerics.Statistics;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using C = DECS_Excel_Add_Ins.Census;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{    /**
     * @brief Are we using full address or just the zip code?
     */
    internal enum LocationSource
    {
        [Description("Address")]
        Address = 1,
        [Description("Zip")]
        Zip = 2,
        [Description("Unknown")]
        Unknown = 0,
    }

    internal class AddressToCensusTract
    {
        private Application application;
        private const int HALFWAY_DOWN_THE_SHEET = 12;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal AddressToCensusTract()
        {
            application = Globals.ThisAddIn.Application;
        }

        private void BuildHistogram(List<ulong> fipsList)
        {
            // Create & set up histogram sheet.
            Worksheet censusHistogramSheet = Utilities.CreateNewNamedSheet("Census Tract Histogram");
            Range target = (Range)censusHistogramSheet.Cells[1, 1];
            target.Value = "Census tract number";
            target.Offset[0, 1].Value = "Number of occurrences";

            Dictionary<ulong, int> censusTractHistogram = new Dictionary<ulong, int>();

            foreach (ulong fips in fipsList)
            {
                if (censusTractHistogram.ContainsKey(fips))
                {
                    censusTractHistogram[fips]++;
                }
                else
                {
                    censusTractHistogram[fips] = 1;
                }
            }

            // Sort by value descending, then by key ascending for tie-breaking.
            var sortedItems = censusTractHistogram.OrderByDescending(pair => pair.Value)
                                  .ThenBy(pair => pair.Key);
            int rowOffset = 1;

            foreach (var item in sortedItems)
            {
                target.Offset[rowOffset, 0].Value = item.Key;
                target.Offset[rowOffset, 1].Value = item.Value;
                rowOffset++;
            }
        }

        internal void Convert(Worksheet worksheet)
        {
            // 1) Find the address column.
            Range locationColumn = FindNamedColumn(worksheet, "address");
            LocationSource locationSource = LocationSource.Unknown;
            ZipCodeConverter zipCodeConverter = null;
            Geocode geocoder = null;

            if (locationColumn == null)
            {
                locationColumn = FindNamedColumn(worksheet, "zip");

                if (locationColumn == null)
                {
                    Utilities.WarnColumnNotFound("address or zip");
                    return;
                }

                locationSource = LocationSource.Zip;
                application.StatusBar = "Reading zip code table.";
                zipCodeConverter = new ZipCodeConverter();
            }
            else
            {
                locationSource = LocationSource.Address;
                application.StatusBar = "Creating census Geocoder object.";
                geocoder = new Geocode();
            }
            Range censusColumn = null;

            if (locationSource == LocationSource.Address)
            {
                censusColumn = Utilities.InsertNewColumn(range: locationColumn, newColumnName: "Census FIPS", side: InsertSide.Right);
            }

            int rowOffset = 1;
            int numConsecutiveFailures = 0;
            List<ulong> fipsAll = new List<ulong>();

            // 2) Convert each address to census tract FIPS number.
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
                        if (locationSource == LocationSource.Address)
                        {
                            C.CensusData data = geocoder.Convert(location);
                            ulong fips = data.FIPS();
                            censusColumn.Offset[rowOffset, 0].Value2 = fips;
                            fipsAll.Add(fips);

                            // reset
                            numConsecutiveFailures = 0;
                        }
                    }
                }
                catch (NullReferenceException)
                {
                    numConsecutiveFailures++;
                }

                // An occasional miss is ok, but three in a row & we've run outta data.
                if (numConsecutiveFailures >= 3)
                {
                    break;
                }

                rowOffset++;
                Utilities.ScrollToRow(worksheet, censusColumn.Offset[rowOffset].Row - HALFWAY_DOWN_THE_SHEET);

                if (rowOffset % 10 == 0)
                {
                    application.StatusBar = "Processed " + rowOffset.ToString() + " addresses.";
                }
            }

            BuildHistogram(fipsAll);
            application.StatusBar = "Complete";
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
    }
}
