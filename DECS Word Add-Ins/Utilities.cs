using DecsWordAddIns.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
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

            // Double up apostrophes so SQL correctly interprets them.
            column_name = column_name.Replace(@"'", "''");

            // Shift To Title Case.
            var textinfo = CultureInfo.CurrentCulture.TextInfo;
            return textinfo.ToTitleCase(column_name);
        }

        internal static string GetUserName()
        {
            string fullUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            // Strip off any prefix like "AD".
            string userName = Path.GetFileName(fullUserName);

            return userName;
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
            // https://stackoverflow.com/a/19329123/18749636
            catch (Exception ex) when (
                ex is System.IO.PathTooLongException
                || ex is System.NotSupportedException)
            {
                output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: true);
                writer_obj = new StreamWriter(output_filename);
            }

            return (writer: writer_obj, opened_filename: output_filename);
        }

        // From a text file, build a dictionary of login names, nice names.
        // gwashington, George Washington
        internal static Dictionary<string,string> ReadUserNamesFile()
        {
            Dictionary<string,string> userNames = new Dictionary<string,string>();
            string allUserNames = Resources.usernames;
            string[] lines = allUserNames.Split('\n');

            foreach (var line in lines)
            {
                string[] pieces = line.Split(':');

                if (pieces.Length == 2)
                {
                    userNames.Add(pieces[0].Trim(), pieces[1].Trim());
                }
            }

            return userNames;
        }

        internal static string SalutationFromName(string name)
        {
            bool isProfessor = name.Contains("Prof");
            bool isPhysician = name.Contains("Dr") || name.Contains("MD");

            string name_detitled = name.Replace("Dr", "");
            name_detitled = name_detitled.Replace("Professor", "");
            name_detitled = name_detitled.Replace("Prof", "");
            name_detitled = name_detitled.Replace("Assistant", "");
            name_detitled = name_detitled.Replace("Asst", "");
            name_detitled = name_detitled.Replace("Associate", "");
            name_detitled = name_detitled.Replace("Assoc", "");
            name_detitled = name_detitled.Replace("MD", "");
            name_detitled = name_detitled.Replace("DO", "");
            name_detitled = name_detitled.Replace("PhD", "");
            name_detitled = name_detitled.Replace("Ph.D.", "");
            name_detitled = name_detitled.Replace("PHD", "");
            string name_depunctuated = name_detitled.Replace(",", "");
            name_depunctuated = name_depunctuated.Trim();
            string[] name_pieces = name_depunctuated.Split(' ');
            string last_name = name_pieces.Last();

            if (isPhysician)
            {
                return "Dr. " + last_name;
            }

            if (isProfessor)
            {
                return "Prof. " + last_name;
            }

            return name_depunctuated;
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
        internal static string TranslateLoginName(string loginName)
        {
            Dictionary<string, string> userNamesList = ReadUserNamesFile();

            if (userNamesList != null && userNamesList.ContainsKey(loginName))
            {
                return userNamesList[loginName];
            }

            return loginName;
        }
    }
}