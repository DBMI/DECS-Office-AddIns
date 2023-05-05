using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        internal static string FormOutputFilename(string filename, string filetype = ".sql", bool short_version = false)
        {
            string dir = Path.GetDirectoryName(filename);
            string just_the_filename = Path.GetFileNameWithoutExtension(filename);
            string sql_filename = Path.Combine(dir, just_the_filename + filetype);

            if (short_version)
            {
                sql_filename = just_the_filename + filetype;
            }
            
            return sql_filename;
        }

        // Detects when a name (which should be like "Hypertension")
        // is actually just a list of codes ("J42", "J43", "J44").
        internal static bool IsJustListOfCodes(string name, MatchCollection matches)
        {
            name = name.Trim();

            // Strip out all the codes & see if anything's left.
            foreach (Match match in matches)
            {
                string code_value = match.Groups[0].Value;

                if (code_value == null) continue;

                name = name.Replace(code_value, "");
            }

            // Remove commas, spaces.
            name = name.Replace(",", "");
            name = name.Replace(" ", "");

            return name.Length == 0;
        }

        // Open the output StreamWriter object,
        // understanding that we might have to substitute a shorter version of the output filename
        // if the default filename is too long.
        internal static (StreamWriter writer, string opened_filename) OpenOutput(string input_filename, string filetype = ".sql")
        {
            string output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: false);
            StreamWriter writer_obj;

            try
            {
                writer_obj = new StreamWriter(output_filename);
            }
            catch (System.IO.PathTooLongException)
            {
                output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: true);
                writer_obj = new StreamWriter(output_filename);
            }

            return (writer: writer_obj, opened_filename: output_filename);
        }

        // Reassure the user that we've created the desired output file,
        // and display the file once they've seen the message.
        internal static void ShowResults(string output_filename)
        {
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            string message = "Created file '" + output_filename + "'.";
            DialogResult result = MessageBox.Show(message, "Success", buttons);

            if (result == DialogResult.OK)
            {
                Process.Start(output_filename);
            }
        }
    }
}