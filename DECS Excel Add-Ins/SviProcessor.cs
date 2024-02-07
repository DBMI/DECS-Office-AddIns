using DECS_Excel_Add_Ins.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    /**
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

    /**
     * @brief Main class for @c AddSVI tool.
     */
    internal class SviProcessor
    {
        private Application application;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal SviProcessor()
        {
            application = Globals.ThisAddIn.Application;
        }

        /// <summary>
        /// Finds either the user-selected column or (if none selected) column with name we expect.
        /// <summary>
        /// <param name="worksheet">Reference to the ActiveSheet.</param>
        /// <param name="lastRowNumber">Number of last row containing data.</param>
        /// <param name="desiredName">Name of column we're looking for.</param>
        /// <returns>Range</returns>
        private Range FindNamedColumn(Worksheet worksheet, int lastRowNumber, string desiredName)
        {
            Regex desiredPattern = new Regex(desiredName.ToLower());
            Range selectedColumn = Utilities.GetSelectedCol(application, lastRowNumber);

            // If user didn't select a column, find it by name.
            if (selectedColumn == null)
            {
                Dictionary<string, Range> columns = Utilities.GetColumnNamesDictionary(worksheet);

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
            // We'll use this in a lot of places, so let's just look it up once.
            int lastRowNumber = Utilities.FindLastRow(worksheet);

            // 1) Find the address column (or the zip column, as a fallback).
            Range locationColumn = FindNamedColumn(worksheet, lastRowNumber, "address");
            LocationSource locationSource = LocationSource.Unknown;
            ZipCodeConverter zipCodeConverter = null;
            Geocode geocoder = null;

            if (locationColumn == null)
            {
                locationColumn = FindNamedColumn(worksheet, lastRowNumber, "zip");

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

            // 2) Populate the SVI dictionary from data file.
            application.StatusBar = "Building SVI table.";
            SviTable sviTable = new SviTable();

            if (sviTable.ready)
            {
                // Build output columns.
                Range sviRankColumn = Utilities.InsertNewColumn(locationColumn, "SVI rank");
                Range sviScoreColumn = Utilities.InsertNewColumn(locationColumn, "SVI score");
                Range censusColumn = null;

                if (locationSource == LocationSource.Address)
                {
                    censusColumn = Utilities.InsertNewColumn(locationColumn, "Census FIPS");
                }

                List<ulong> fipsList;

                // 3) Convert each address or zip to census tract FIPS number, then lookup SVI.
                for (int rowOffset = 1; rowOffset <= lastRowNumber; rowOffset++)
                {
                    try
                    {
                        string location = locationColumn.Offset[rowOffset, 0].Text;

                        if (!string.IsNullOrEmpty(location))
                        {
                            if (locationSource == LocationSource.Address)
                            {
                                CensusData data = geocoder.Convert(location);
                                ulong fips = data.FIPS();
                                censusColumn.Offset[rowOffset, 0].Value2 = fips;
                                fipsList = new List<ulong>();
                                fipsList.Add(fips);
                            }
                            else
                            {
                                fipsList = zipCodeConverter.Convert(location);
                            }

                            // Don't display nonsense numbers (represented by -1).

                            double rawScore = sviTable.raw(fipsList);

                            if (rawScore >= 0)
                            {
                                sviScoreColumn.Offset[rowOffset, 0].Value2 = rawScore;
                            }

                            double rank = sviTable.rank(fipsList);

                            if (rank >= 0)
                            {
                                sviRankColumn.Offset[rowOffset, 0].Value2 = rank;
                            }
                        }
                    }
                    catch
                    {
                        break;
                    }

                    if (rowOffset % 100 == 0)
                    {
                        application.StatusBar = "Processed " + rowOffset.ToString() + "/" + lastRowNumber.ToString() + " patients.";
                    }
                }

                application.StatusBar = "Complete";
            }
        }
    }
}
