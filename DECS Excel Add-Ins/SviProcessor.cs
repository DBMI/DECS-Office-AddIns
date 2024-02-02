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
    internal enum LocationSource
    {
        [Description("Address")]
        Address = 1,
        [Description("Zip")]
        Zip = 2,
        [Description("Unknown")]
        Unknown = 0,
    }

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

        // Scans the worksheet:
        // 1) Finds the address column (or the zip column, if address not found),
        //    either using the selected column or finding it by name.
        // 2) Reads data file California.csv & populates a dictionary mapping census tract # to SVI values.
        // 3) Uses online geocode service to retrieve the census tract for each address.
        // 4) Looks up the SVI values from the tract dictionary.
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
                Range censusColumn = Utilities.InsertNewColumn(locationColumn, "Census FIPS");
                ulong fips;

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
                                fips = data.FIPS();
                            }
                            else
                            {
                                fips = zipCodeConverter.Convert(location);
                            }

                            censusColumn.Offset[rowOffset, 0].Value2 = fips;
                            sviScoreColumn.Offset[rowOffset, 0].Value2 = sviTable.raw(fips);
                            sviRankColumn.Offset[rowOffset, 0].Value2 = sviTable.rank(fips);
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
