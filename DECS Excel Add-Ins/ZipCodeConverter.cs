using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Builds a @c Dictionary<string, List<ulong>> from the nationwide zip code-->census tract file.
     * File available from https://www.huduser.gov/portal/datasets/usps_crosswalk.html
     */
    internal class ZipCodeConverter
    {
        private Dictionary<string, List<ulong>> zipToTractTable;
        internal bool ready { get; } = false;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        /// <summary>
        /// Reads the zip crosswalk file & creates @c Dictionary to map each zip code to a list of census tract numbers.
        /// </summary>
        internal ZipCodeConverter()
        {
            zipToTractTable = new Dictionary<string, List<ulong>>();

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith("ZIP_TRACT_122023.csv"));

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream(resourceName)))
            {
                string[] lines = reader.ReadToEnd().Split('\n');

                // Find ZIP and TRACT in the first row.
                string[] headers = lines[0].Split(',');
                int ZIP_index = Array.IndexOf(headers, "ZIP");

                if (ZIP_index < 0)
                {
                    log.Error("Unable to find 'ZIP' in header.");
                    return;
                }

                int TRACT_index = Array.IndexOf(headers, "TRACT");

                if (TRACT_index < 0)
                {
                    log.Error("Unable to find 'TRACT' in header.");
                    return;
                }

                int maxIndex = Math.Max(ZIP_index, TRACT_index);
                string zip;
                List<ulong> fipsList = new List<ulong>();

                foreach (string line in lines.Skip(1))
                {
                    string[] pieces = line.Split(',');

                    if (pieces.Length > maxIndex)
                    {
                        if (ulong.TryParse(pieces[TRACT_index], out ulong fips))
                        {
                            zip = pieces[ZIP_index];

                            // Force it to have five digits.
                            zip = zip.PadLeft(5, '0');

                            if (zipToTractTable.ContainsKey(zip))
                            {
                                // Get the existing list of FIPS codes for this zip code.
                                fipsList = zipToTractTable[zip];
                            }
                            else
                            {
                                // Create a new blank list;
                                fipsList = new List<ulong>();
                            }

                            // Tack this FIPS code onto the list...
                            fipsList.Add(fips);

                            // ...and replace the dictionary value for this zip code.
                            zipToTractTable[zip] = fipsList;
                        }
                    }
                }

                ready = zipToTractTable.Count > 0;
            }
        }

        /// <summary>
        /// Looks up the zip code in the @c Dictionary
        /// </summary>
        /// <param name="zip">zip code as string</param>
        /// <returns>List<ulong></returns>
        internal List<ulong> Convert(string zip)
        {
            // Strip out any other text (like "NJ 07003").
            zip = Regex.Replace(zip, @"\D", "");
            zip = zip.Trim();

            // Limit zip string to be five digits to be compatible with lookup table.
            if (zip.Length > 5)
            {
                zip = zip.Substring(0, 5);
            }

            // Force it to have five digits.
            zip = zip.PadLeft(5, '0');

            if (ready && zipToTractTable.ContainsKey(zip))
            {
                return zipToTractTable[zip];
            }

            return new List<ulong>(0);
        }
    }
}
