using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DecsWordAddIns
{
    internal static class Utilities
    {
        // Convert fancy Windows stuff so later Regexs work more simply.
        internal static string CleanText(string text) 
        {
            string text_cleaned = text.Trim();
            text_cleaned = text_cleaned.Replace(@"–", "-"); // Replace Windows dash with simple hyphen.
            text_cleaned = text_cleaned.Replace(@"’", "'"); // Replace Windows apostrophe with simple apostrophe.

            return text_cleaned;
        }

        // Turn a free-text description of a condition (like "Interstitial lung disease")
        // into a string suitable for use as a SQL column name ("Interstitial_Lung_Disease").
        internal static string CleanNameForSql(string condition_name)
        {
            string column_name = condition_name.Trim().Replace(' ', '_');
            column_name = column_name.Replace(',', '_');
            column_name = column_name.Replace("__", "_");
            var textinfo = CultureInfo.CurrentCulture.TextInfo;
            return textinfo.ToTitleCase(column_name);
        }

        // Turn the statement of work filename into a .sql filename.
        internal static string FormOutputFilename(string filename, string filetype = ".sql")
        {
            string dir = Path.GetDirectoryName(filename);
            string just_the_filename = Path.GetFileNameWithoutExtension(filename);
            string sql_filename = Path.Combine(dir, just_the_filename + filetype);
            return sql_filename;
        }
    }
}