using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public enum SviScope
    {
        [Description("California")]
        California = 1,
        [Description("USA")]
        USA = 2,
        [Description("Unknown")]
        Unknown = 0,
    }

    /**
     * @brief Holds both the raw score (from "SPL_THEMES" column) and fractional ranking ("RPL_THEMES").
     */
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

    /**
     * @brief Builds & uses a Dictionary that maps the FIPS code to the SVI information.
     */
    internal class SviTable
    {
        private Dictionary<ulong, SviScore> sviTable;
        private string filePattern;
        private char fileSeparator;

        internal bool ready { get; } = false;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            MethodBase.GetCurrentMethod().DeclaringType
        );

        internal SviTable()
        {
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;

            sviTable = new Dictionary<ulong, SviScore>();

            var assembly = Assembly.GetExecutingAssembly();

            GetSviFileNameAndSeparator();

            if (string.IsNullOrEmpty(filePattern))
            {
                return;
            }

            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith(filePattern));

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream(resourceName)))
            {
                string[] lines = reader.ReadToEnd().Split('\n');

                // Find FIPS and RPL_THEMES in the first row.
                string[] headers = lines[0].Split(fileSeparator);
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
                int numLinesProcessed = 0;
                int numLinesPresent = lines.Length;

                foreach (string line in lines.Skip(1))
                {
                    string[] pieces = line.Split(fileSeparator);

                    if (pieces.Length > maxIndex)
                    {
                        if (ulong.TryParse(pieces[FIPS_index], out ulong fips))
                        {
                            SviScore sviObj = new SviScore(pieces[SPL_THEMES_index], pieces[RPL_THEMES_index]);
                            sviTable.Add(fips, sviObj);
                        }
                    }

                    numLinesProcessed++;

                    if (numLinesProcessed % 100 == 0)
                    {
                        application.StatusBar = "Processed " + numLinesProcessed.ToString() + "/" + numLinesPresent.ToString() + " rows.";
                    }
                }

                application.StatusBar = "Complete";

                ready = sviTable.Count > 0;
            }
        }

        /// <summary>
        /// Gets either the all-USA file or just the California data.
        /// <summary>
        /// <returns>string</returns>
        private string GetSviFileName(SviScope desiredScope)
        {
            string filePattern = string.Empty;

            switch (desiredScope)
            {
                case SviScope.California:
                    filePattern = "California.csv";
                    break;
                case SviScope.USA:
                    filePattern = "SVI_2022_US.csv";
                    break;
            }

            return filePattern;
        }

        /// <summary>
        /// Gets file separator for
        /// the all-USA file (comma-separated) or just the California data (tab-separated).
        /// <summary>
        /// <returns>string</returns>
        private char GetSviFileSeparator(SviScope desiredScope)
        {
            char fileSeparator = '\0';

            switch (desiredScope)
            {
                case SviScope.California:
                    fileSeparator = ',';
                    break;
                case SviScope.USA:
                    fileSeparator = ',';
                    break;
            }

            return fileSeparator;
        }

        /// <summary>
        /// Sets the class properties depending on which data file we want.
        /// <summary>
        /// <returns>SviScope enum</returns>
        private void GetSviFileNameAndSeparator()
        {
            SviScope desiredScope = GetUserPreference();
            filePattern = GetSviFileName(desiredScope);
            fileSeparator = GetSviFileSeparator(desiredScope);
        }

        /// <summary>
        /// Asks the user if they want all-USA file or just California data.
        /// <summary>
        /// <returns>SviScope enum</returns>
        private SviScope GetUserPreference()
        {
            using (UseCalforniaOrAllUsaForm form = new UseCalforniaOrAllUsaForm())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return form.scope;
                }
            }

            return SviScope.Unknown;
        }

        /// <summary>
        /// Returns the average rank across all the census tracts provided.
        /// <summary>
        /// <param name="tractList">List<ulong> census tract numbers.</param>
        /// <returns>double</returns>
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

        /// <summary>
        /// Returns the average raw score across all the census tracts provided.
        /// <summary>
        /// <param name="tractList">List<ulong> census tract numbers.</param>
        /// <returns>double</returns>
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
