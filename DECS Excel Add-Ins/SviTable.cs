using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    // Both the raw score (from "SPL_THEMES" column) and percentile ranking ("RPL_THEMES").
    internal class SviScore
    {
        internal double rawScore { get; }
        internal int percentile { get; }

        internal SviScore(string rawStr, string pctStr)
        {
            if (double.TryParse(rawStr, out double dTemp))
            {
                rawScore = dTemp;
            }

            if (int.TryParse(pctStr.TrimEnd('%'), out int iTemp))
            {
                percentile = iTemp;
            }
        }
    }
    internal class SviTable
    {
        private Dictionary<ulong, SviScore> sviTable;
        internal bool ready { get; } = false;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal SviTable()
        {
            sviTable = new Dictionary<ulong, SviScore>();

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith("California.csv"));

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream(resourceName)))
            {
                string[] lines = reader.ReadToEnd().Split('\n');

                // Find FIPS and RPL_THEMES in the first row.
                string[] headers = lines[0].Split('\t');
                int FIPS_index = Array.IndexOf(headers, "FIPS");

                if (FIPS_index == 0)
                {
                    log.Error("Unable to find 'FIPS' in header.");
                    return;
                }

                int SPL_THEMES_index = Array.IndexOf(headers, "SPL_THEMES");

                if (SPL_THEMES_index == 0)
                {
                    log.Error("Unable to find 'SPL_THEMES' in header.");
                    return;
                }

                int RPL_THEMES_index = Array.IndexOf(headers, "RPL_THEMES");

                if (RPL_THEMES_index == 0)
                {
                    log.Error("Unable to find 'RPL_THEMES' in header.");
                    return;
                }

                int maxIndex = Math.Max(FIPS_index, Math.Max(SPL_THEMES_index, RPL_THEMES_index));

                foreach (string line in lines.Skip(1))
                {
                    string[] pieces = line.Split('\t');

                    if (pieces.Length > maxIndex)
                    {
                        if (ulong.TryParse(pieces[FIPS_index], out ulong fips))
                        {
                            SviScore sviObj = new SviScore(pieces[SPL_THEMES_index], pieces[RPL_THEMES_index]);
                            sviTable.Add(fips, sviObj);
                        }
                    }
                }

                ready = sviTable.Count > 0;
            }
        }

        internal int percentile(ulong tract)
        {
            if (sviTable.ContainsKey(tract))
            {
                return sviTable[tract].percentile;
            }

            return 0;
        }

        internal double raw(ulong tract)
        {
            if (sviTable.ContainsKey(tract))
            {
                return sviTable[tract].rawScore;
            }

            return 0;
        }
    }
}
