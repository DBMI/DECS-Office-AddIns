using DECS_Excel_Add_Ins.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
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

        private Range FindAddressColumn(Worksheet worksheet, int lastRowNumber)
        {
            Range addressColumn = Utilities.GetSelectedCol(application, lastRowNumber);

            // If user didn't select a column, find it by name.
            if (addressColumn is null)
            {
                Regex addressPattern = new Regex(@"address");
                Dictionary<string, Range> columns = Utilities.GetColumnNamesDictionary(worksheet);

                foreach (KeyValuePair<string, Range> column in columns)
                {
                    Match match = addressPattern.Match(column.Key.ToLower());

                    if (match.Success)
                    {
                        addressColumn = column.Value;
                        break;
                    }
                }
            }

            return addressColumn;
        }

        // Scans the worksheet:
        // 1) Finds the address column, either using the selected column or finding it by name.
        // 2) Reads data file California.csv & populates a dictionary mapping census tract # to SVI values.
        // 3) Uses online geocode service to retrieve the census tract for each address.
        // 4) Looks up the SVI values from the tract dictionary.
        internal void Scan(Worksheet worksheet)
        {
            // We'll use this in a lot of places, so let's just look it up once.
            int lastRowNumber = Utilities.FindLastRow(worksheet);

            // 1) Find the address column.
            Range addressColumn = FindAddressColumn(worksheet, lastRowNumber);

            if (addressColumn is null)
            {
                Utilities.WarnColumnNotFound("address");
                return;
            }

            // 2) Populate the SVI dictionary from data file.
            SviTable sviTable = new SviTable();

            if (sviTable.ready)
            {
                // 3) Convert each address to census tract FIPS number, then lookup SVI.
                Range sviPercentileColumn = Utilities.InsertNewColumn(addressColumn, "SVI %");
                Range sviScoreColumn = Utilities.InsertNewColumn(addressColumn, "SVI score");
                Range censusColumn = Utilities.InsertNewColumn(addressColumn, "Census FIPS");
                Geocode geocoder = new Geocode();

                for (int rowOffset = 1; rowOffset <= lastRowNumber; rowOffset++)
                {
                    try
                    {
                        string address = addressColumn.Offset[rowOffset, 0].Value2;

                        if (!string.IsNullOrEmpty(address))
                        {
                            CensusData data = geocoder.Convert(address);
                            ulong fips = data.FIPS();
                            censusColumn.Offset[rowOffset, 0].Value2 = fips;
                            sviPercentileColumn.Offset[rowOffset, 0].Value2 = sviTable.percentile(fips);
                            sviScoreColumn.Offset[rowOffset, 0].Value2 = sviTable.raw(fips);
                        }
                    }
                    catch
                    {
                        break;
                    }
                }
            }

        }
    }
}
