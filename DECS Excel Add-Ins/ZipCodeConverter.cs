using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    internal class ZipCodeConverter
    {
        private Dictionary<string, ulong> zipToTractTable;
        internal bool ready { get; } = false;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal ZipCodeConverter()
        {
            zipToTractTable = new Dictionary<string, ulong>();

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

                foreach (string line in lines.Skip(1))
                {
                    string[] pieces = line.Split(',');

                    if (pieces.Length > maxIndex)
                    {
                        if (ulong.TryParse(pieces[TRACT_index], out ulong fips))
                        {
                            zip = pieces[ZIP_index];
                            
                            if (!zipToTractTable.ContainsKey(zip))
                            {
                                zipToTractTable.Add(zip, fips);
                            }
                        }
                    }
                }

                ready = zipToTractTable.Count > 0;
            }
        }

        internal ulong Convert(string zip)
        {
            if (ready && zipToTractTable.ContainsKey(zip))
            {
                return zipToTractTable[zip];
            }

            return 0;
        }
    }
}
