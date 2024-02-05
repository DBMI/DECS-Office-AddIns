using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    // Both the raw score (from "SPL_THEMES" column) and fractional ranking ("RPL_THEMES").
    internal class SviScore
    {
        internal double rawScore { get; }
        internal double rank { get; }

        internal SviScore(string rawStr, string rankingStr)
        {
            if (double.TryParse(rawStr, out double dScore))
            {
                rawScore = dScore;
            }

            // The California-only table reports ranking in interger percents, but the all-US table uses fractions.
            if (rankingStr.Contains("%"))
            {
                if (int.TryParse(rankingStr.TrimEnd('%'), out int iRank))
                {
                    // Convert 3% to 0.03
                    rank = iRank * 0.01;
                }
            }
            else
            {
                if (double.TryParse(rankingStr, out double dRank))
                {
                    rank = dRank;
                }
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
            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith("SVI_2020_US.csv"));

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream(resourceName)))
            {
                string[] lines = reader.ReadToEnd().Split('\n');

                // Find FIPS and RPL_THEMES in the first row.
                string[] headers = lines[0].Split(',');
                int FIPS_index = Array.IndexOf(headers, "FIPS");

                if (FIPS_index < 0)
                {
                    log.Error("Unable to find 'FIPS' in header.");
                    return;
                }

                int SPL_THEMES_index = Array.IndexOf(headers, "SPL_THEMES");

                if (SPL_THEMES_index < 0)
                {
                    log.Error("Unable to find 'SPL_THEMES' in header.");
                    return;
                }

                int RPL_THEMES_index = Array.IndexOf(headers, "RPL_THEMES");

                if (RPL_THEMES_index < 0)
                {
                    log.Error("Unable to find 'RPL_THEMES' in header.");
                    return;
                }

                int maxIndex = Math.Max(FIPS_index, Math.Max(SPL_THEMES_index, RPL_THEMES_index));

                foreach (string line in lines.Skip(1))
                {
                    string[] pieces = line.Split(',');

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

        internal double rank(List<ulong> tractList)
        {
            double sum = 0;
            int numValues = 0;

            foreach (ulong tract in tractList)
            {
                // Don't average in N/A values like -999.
                if (sviTable.ContainsKey(tract) && sviTable[tract].rank >= 0)
                {
                    sum += sviTable[tract].rank;
                    numValues++;
                }                
            }

            if (numValues > 0)
            {
                return sum / numValues;
            }

            return -1.0;
        }

        internal double raw(List<ulong> tractList)
        {
            double sum = 0;
            int numValues = 0;

            foreach (ulong tract in tractList)
            {
                // Don't average in N/A values like -999.
                if (sviTable.ContainsKey(tract) && sviTable[tract].rawScore >= 0)
                {
                    sum += sviTable[tract].rawScore;
                    numValues++;
                }
            }

            if (numValues > 0)
            {
                return sum / numValues;
            }

            return -1.0;
        }
    }
}
