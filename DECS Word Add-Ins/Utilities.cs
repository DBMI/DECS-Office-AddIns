using DecsWordAddIns.Properties;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace DecsWordAddIns
{
    /// <summary>
    /// Useful methods.
    /// </summary>
    internal static class Utilities
    {
        /// <summary>
        /// Pull together ALL the Document's text.
        /// </summary>
        /// <param name="doc">Word @c Document object</param>
        /// <returns>string</returns>
        internal static string AllText(Document doc)
        {
            string combinedText = string.Empty;

            // Splice together text from all paragraphs.
            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                combinedText += paragraph.Range.Text;
            }

            // Remove "/a" `alert` characters.
            char alert = '\u0007';
            return combinedText.Replace(alert.ToString(), "");
        }

        /// <summary>
        /// Convert/remove stuff so later Regexs don't have to allow for them.
        /// </summary>
        /// <param name="text">input text</param>
        /// <returns>string</returns>
        internal static string CleanText(string text)
        {
            string text_cleaned = text.Trim();
            text_cleaned = text_cleaned.Replace(@"–", "-"); // Replace Windows dash with simple hyphen.
            text_cleaned = text_cleaned.Replace(@"’", "'"); // Replace Windows apostrophe with simple apostrophe.
            char verticalTab = '\u000B';
            text_cleaned = text_cleaned.Replace(verticalTab.ToString(), " "); // Remove \v chars.
            text_cleaned = text_cleaned.Replace(Environment.NewLine, "");
            return text_cleaned;
        }

        /// <summary>
        /// Turn a free-text description of a condition (like "Interstitial lung disease")
        /// into a string suitable for use as a SQL column name ("Interstitial_Lung_Disease").
        /// </summary>
        /// <param name="condition_name">string</param>
        /// <returns>string</returns>
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

        /// <summary>
        /// Uses system utility to get the name of the user.
        /// </summary>
        /// <returns>string</returns>
        internal static string GetUserName()
        {
            string fullUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            // Strip off any prefix like "AD".
            string userName = Path.GetFileName(fullUserName);

            return userName;
        }

        /// <summary>
        /// Turn the statement of work filename into a .sql filename.
        /// </summary>
        /// <param name="filename">SoW filename</param>
        /// <param name="filetype">type of file to create (default = ".sql")</param>
        /// <param name="short_version">bool: Just filename.type? (default = false)</param>
        /// <returns></returns>
        internal static string FormOutputFilename(
            string filename,
            string filetype = ".sql",
            bool short_version = false
        )
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

        /// <summary>
        /// Detects when a name (which should be like "Hypertension")
        /// is actually just a list of codes ("J42", "J43", "J44").
        /// </summary>
        /// <param name="name">string to parse</param>
        /// <param name="matches">@c MatchColllection object</param>
        /// <returns>bool</returns>
        internal static bool IsJustListOfCodes(string name, MatchCollection matches)
        {
            name = name.Trim();

            // Strip out all the codes & see if anything's left.
            foreach (Match match in matches)
            {
                string code_value = match.Groups[0].Value;

                if (code_value == null)
                    continue;

                name = name.Replace(code_value, "");
            }

            // Remove commas, spaces.
            name = name.Replace(",", "");
            name = name.Replace(" ", "");

            return name.Length == 0;
        }

        /// <summary>
        /// Open the output StreamWriter object,
        /// understanding that we might have to substitute a shorter version of the output filename
        /// if the default filename is too long.
        /// </summary>
        /// <param name="input_filename">SoW filename</param>
        /// <param name="filetype">type of file to create (default = ".sql")</param>
        /// <returns>Tuple(StreamWriter, string)</returns>
        internal static (StreamWriter writer, string openedFilename) OpenOutput(
            string input_filename,
            string filetype = ".sql"
        )
        {
            string outputFilename = Utilities.FormOutputFilename(
                filename: input_filename,
                filetype: filetype,
                short_version: false
            );
            StreamWriter writer_obj;

            try
            {
                writer_obj = new StreamWriter(outputFilename);
            }
            // https://stackoverflow.com/a/19329123/18749636
            catch (Exception ex)
                when (ex is System.IO.PathTooLongException || ex is System.NotSupportedException)
            {
                outputFilename = Utilities.FormOutputFilename(
                    filename: input_filename,
                    filetype: filetype,
                    short_version: true
                );
                writer_obj = new StreamWriter(outputFilename);
            }

            return (writer: writer_obj, openedFilename: outputFilename);
        }

        /// <summary>
        /// If condition name isn't empty, return it prepended with "---".
        /// </summary>
        /// <param name="conditionName">string to process</param>
        /// <param name="separator">string between previous text & hyphens (default = "")</param>
        /// <returns>string</returns>
        internal static string PrependWithHypens(string conditionName, string separator = "")
        {
            string conditionNameWithHyphens = "";

            // Prepend with "---".
            if (conditionName.Length > 0)
            {
                conditionNameWithHyphens = separator + " --- " + conditionName;
            }

            return conditionNameWithHyphens;
        }

        /// <summary>
        /// From a text file, build a dictionary of login names, nice names.
        /// Example: gwashington, George Washington
        /// </summary>
        /// <returns>Dictionary<string, string></returns>
        internal static Dictionary<string, string> ReadUserNamesFile()
        {
            Dictionary<string, string> userNames = new Dictionary<string, string>();
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

        /// <summary>
        /// Run a string replacement again & again until it stops matching.
        /// Can be either straight string @c .Replace or Regex.
        /// </summary>
        /// <param name="text">string to process</param>
        /// <param name="pattern">string: Regex pattern</param>
        /// <param name="replacement">string: what to replace matches with</param>
        /// <param name="isRegex">bool: is this a Regex? (default = false)</param>
        /// <returns>string</returns>
        internal static string ReplaceUntilNoMore(string text, string pattern, string replacement, bool isRegex = false)
        {
            string textCleaned = text;
            long cleanedLength = long.MaxValue;

            while (textCleaned.Length < cleanedLength)
            {
                cleanedLength = textCleaned.Length;

                if (isRegex)
                {
                    textCleaned = Regex.Replace(textCleaned, pattern, replacement);
                }
                else
                {
                    textCleaned = textCleaned.Replace(pattern, replacement);
                }
            }

            return textCleaned;
        }

        /// <summary>
        /// Turn a name like "Prof. Able Baker Charlie, MD" into "Prof. Charlie"
        /// </summary>
        /// <param name="name">string to parse</param>
        /// <returns>string</returns>
        internal static string SalutationFromName(string name)
        {
            bool isProfessor = name.Contains("Prof");
            bool isPharmacist = name.Contains("Pharm.D.");
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
            name_detitled = name_detitled.Replace("PharmD", "");
            name_detitled = name_detitled.Replace("Pharm.D", "");
            name_detitled = name_detitled.Replace("Pharm.D.", "");
            name_detitled = name_detitled.Replace("PharmD.", "");
            name_detitled = name_detitled.Replace("PhD", "");
            name_detitled = name_detitled.Replace("Ph.D.", "");
            name_detitled = name_detitled.Replace("PHD", "");
            string name_depunctuated = name_detitled.Replace(",", "");
            name_depunctuated = name_depunctuated.Trim();
            string[] name_pieces = name_depunctuated.Split(' ');
            string last_name = name_pieces.Last();

            if (isPhysician || isPharmacist)
            {
                return "Dr. " + last_name;
            }

            if (isProfessor)
            {
                return "Prof. " + last_name;
            }

            return name_depunctuated;
        }

        /// <summary>
        /// Returns the user-selected portion of the document.
        /// </summary>
        /// <param name="doc">Word @c Document object</param>
        /// <returns>List<string></returns>
        // https://stackoverflow.com/a/11025539/18749636
        internal static List<string> SelectedText(Document doc)
        {
            List<string> textBlocks = new List<string>();

            Application app = doc.Application;
            Selection wordSelection = app.Selection;
            string allText;

            if (wordSelection != null &&
                wordSelection.Text != null &&
                wordSelection.Text.Length > 10)
            {
                allText = wordSelection.Text;
            }
            else
            {
                // Ignore how Word defines "Paragraphs".
                // We'll concatentate all the text...
                allText = AllText(doc);
            }

            // ...then split it how WE define paragraphs.
            textBlocks = SplitAtBlankLines(allText);
            return textBlocks;
        }

        /// <summary>
        /// Splits text into List<string>, split at blank lines.
        /// </summary>
        /// <param name="text">string to parse</param>
        /// <returns>List<string></returns>
        private static List<string> SplitAtBlankLines(string text)
        {
            List<string> textBlocks = new List<string>();
            
            if (text == null || string.IsNullOrEmpty(text))
            {
                return textBlocks;
            }

            char carriageReturn = '\u000D';
            string cr = carriageReturn.ToString();

            // Replace /r/r/r/r with just /r/r.
            string fourCarriageReturns = cr + cr + cr + cr;
            string doubleCarriageReturns = cr + cr;
            string textCleaned = ReplaceUntilNoMore(text, fourCarriageReturns, doubleCarriageReturns);

            // https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/strings/
            // Split at EITHER double NewLines OR double VerticalTabs OR double carriage returns.
            char verticalTab = '\u000B';
            string vt = verticalTab.ToString();
            string[] doubleReturns = new string[] { vt + vt,
                                                    cr + cr,
                                                    Environment.NewLine + Environment.NewLine};
            textBlocks = textCleaned.Split(doubleReturns, StringSplitOptions.RemoveEmptyEntries).ToList();
            return textBlocks;
        }

        //internal static string TranslateLoginName(string loginName)
        //{
        //    Dictionary<string, string> userNamesList = ReadUserNamesFile();

        //    if (userNamesList != null && userNamesList.ContainsKey(loginName))
        //    {
        //        return userNamesList[loginName];
        //    }

        //    return loginName;
        //}
    }
}
