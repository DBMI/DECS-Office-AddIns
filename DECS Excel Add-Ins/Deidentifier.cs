using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI;
using System.Windows.Forms;
using System.Xml.Linq;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    // What did the user select in the HideThisNameForm?
    internal class SelectionResult
    {
        internal string alias;
        internal bool replace;
        internal string wordToReplace;

        internal SelectionResult(string alias, bool replace, string wordToReplace)
        {
            this.alias = alias;
            this.replace = replace;
            this.wordToReplace = wordToReplace;
        }
    }

    internal class Deidentifier
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private bool cancel = false;
        
        private const string dateOnlyPattern = @"\d{1,2}\/\d{1,2}\/\d{4}[\s\.](?!\d)";
        private Regex dateOnlyRegex;
        private const string dateTimePattern = @"\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s[AP]M";
        private Regex dateTimeRegex;
        private const string drNamePattern = @"Dr\.\s[A-Z]\w+,?\s*(?:[A-Z]\.\s*)?(?:[A-Z]\w+)?";
        private Regex drNameRegex;
        private const string monthDayOnlyPattern = @"\d{1,2}\/\d{1,2}(?![\/\d])";
        private Regex monthDayOnlyRegex;
        private const string monthSpelledOutPattern = @"\w{4,8}\s*\d{4}";
        private Regex monthSpelledOutRegex;
        private const string namePattern = @"[A-Z][a-z]+";
        private Regex nameRegex;
        private const string providerNameTitlePattern = @"[A-Z][\w-]+,?\s*[A-Z][\w-]*\.?\s*(?:[A-Z][\w-]*\.?)?,?\s*[A-Z]{2,}(?:\s[A-Z-]{2,})?";
        private Regex providerNameTitleRegex;
        private const string wordsBeforePattern = @"(?<left>(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?)";
        private const string wordsAfterPattern = @"(?<right>(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?)";

        private int dayOffset;
        private int monthOffset;
        private TimeSpan deltaT;
        
        private int lastRow;
        
        private Range selectedColumnRng;
        private List<Range> selectedColumnsRng;

        private Dictionary<string, string> namesAndAliasesDict;
        private List<string> namesToSkip;
        private string[] workInProgress;

        internal Deidentifier()
        {
            application = Globals.ThisAddIn.Application;
        }

        // https://learn.microsoft.com/en-us/troubleshoot/developer/visualstudio/csharp/language-compilers/compute-hash-values
        private static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);

            for (i = 0; i < arrInput.Length; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }

            return sOutput.ToString();
        }

        private SelectionResult EncodeProviderName(string nameString,
                                          string leftContext,
                                          string rightContext)
        {
            string nameCleaned = nameString.Trim();
            string alias = nameCleaned;
            bool replace = false;

            if (namesAndAliasesDict.ContainsKey(nameCleaned))
            {
                alias = namesAndAliasesDict[nameCleaned];
                replace = true;
            }
            else if (!namesToSkip.Contains(nameCleaned))
            {
                List<string> similarNames = FindSimilarNames(nameString);

                using (HideThisNameForm form = new HideThisNameForm(nameCleaned, 
                                                                    similarNames, 
                                                                    leftContext, 
                                                                    rightContext,
                                                                    FindSimilarNames))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.Cancel)
                    {
                        cancel = true;
                    }
                    else if(result == DialogResult.OK)
                    {
                        // What word are we replacing?
                        if (!string.IsNullOrEmpty(form.chosenName))
                        {
                            // This is a NEW name to be hidden (not the one we sent to the form).
                            nameCleaned = form.chosenName.Trim();
                        }

                        if (string.IsNullOrEmpty(form.linkedName))
                        {
                            // This is a NEW name to be hidden.
                            string hashCode = String.Format("{0:X}", nameCleaned.GetHashCode());
                            alias = "<" + hashCode + ">";
                            namesAndAliasesDict.Add(nameCleaned, alias);
                        }
                        else
                        {
                            // This is a reference to an existing name.
                            alias = namesAndAliasesDict[form.linkedName];

                            // Perhaps we want to link a new name to an existing alias.
                            // Example: Already have an entry for "Dr. Able Provider" and we
                            // want the SAME alias for "Provider, Able, MD".
                            // (This includes the case where we've edited the string presented
                            // via form.chosenName.)
                            if (!namesAndAliasesDict.ContainsKey(nameCleaned))
                            {
                                namesAndAliasesDict[nameCleaned] = alias;
                            }
                        }

                        replace = true;
                    }
                    else 
                    {
                        namesToSkip.Add(nameCleaned);
                    }
                }
            }

            return new SelectionResult(alias: alias, replace: replace, wordToReplace: nameCleaned);
        }

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application, lastRow);

            if (selectedColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        selectedColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        private bool FindSelectedColumns(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnsRng = Utilities.GetSelectedCols(application, lastRow);

            if (selectedColumnsRng is null)
            {
                // Then ask user to select columns of interest.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: true))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        foreach (string selectedColumnName in form.selectedColumns)
                        {
                            Range thisRng = Utilities.TopOfNamedColumn(worksheet, selectedColumnName);
                            selectedColumnsRng.Add(thisRng);
                        }

                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        private List<string> FindSimilarNames(string name = "")
        {
            List<string> similarNames = new List<string>();

            if (string.IsNullOrEmpty(name))
            {
                // It's a signal to send ALL the keys.
                similarNames = namesAndAliasesDict.Keys.ToList<string>();
            }
            else
            {
                double editDistanceThreshold = 0.5;
                double fractionOfWordsPresentThreshold = 0.5;

                Fastenshtein.Levenshtein lev = new Fastenshtein.Levenshtein(name);

                foreach (string key in namesAndAliasesDict.Keys.ToList<string>())
                {
                    // Is every word in this new name present in an existing key?
                    if (Utilities.WordsPresent(name, key) >= fractionOfWordsPresentThreshold)
                    {
                        similarNames.Add(key);
                        continue;
                    }

                    // Test using Levenshtein distance.
                    double wordLength = (double)Math.Min(name.Length, key.Length);
                    int levenshteinDistance = lev.DistanceFrom(key);
                    double relativeDistance = levenshteinDistance / wordLength;

                    if (relativeDistance <= editDistanceThreshold)
                    {
                        similarNames.Add(key);
                    }
                }
            }

            similarNames.Sort();
            return similarNames;
        }

        internal void GenerateHash(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumns(worksheet))
            {
                // Make room for new column.
                Range hashColumn = Utilities.InsertNewColumn(range: selectedColumnsRng.Last(), newColumnName: "Scrambled Identifier", side: InsertSide.Right);

                string sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, hashColumn.Column];
                    sourceData = Utilities.CombineColumns(worksheet, rowNumber, selectedColumnsRng);

                    if (!string.IsNullOrEmpty(sourceData))
                    {
                        target.Value = StringToHash(sourceData);
                    }
                }
            }
        }

        internal void HidePhysicianNames(Worksheet worksheet)
        {
            // Initialize needed variables.
            lastRow = worksheet.UsedRange.Rows.Count;
            namesAndAliasesDict = new Dictionary<string, string>();
            namesToSkip = new List<string>();

            // Instantiate reusable Regexes.
            drNameRegex = new Regex(drNamePattern);
            nameRegex = new Regex(namePattern);
            providerNameTitleRegex = new Regex(providerNameTitlePattern);

            if (FindSelectedColumns(worksheet))
            {
                foreach(Range col in selectedColumnsRng)
                {
                    HidePhysicianNamesOneColumn(col);
                }
            }
        }

        private void HidePhysicianNamesOneColumn(Range selectedCol)
        {
            string selectedColumnName = selectedCol.Value.ToString();
            string newColumnName = selectedColumnName + " (Names Hidden)";

            // Clear out the buffer for the new column.
            workInProgress = new string[lastRow + 1];

            // Make room for new column.
            Range aliasedColumn = Utilities.InsertNewColumn(range: selectedCol,
                                                            newColumnName: newColumnName,
                                                            side: InsertSide.Right);

            string sourceData;
            Range target;
            Worksheet worksheet = selectedCol.Worksheet;

            // Run the more detailed rules first to assemble a library of names
            // so can later match "Able" with "Dr. Able Provider".
            for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
            {
                if (cancel) { break; }

                sourceData = worksheet.Cells[rowNumber, selectedCol.Column].Value;

                // Modify & store as workInProgress.
                if (!string.IsNullOrEmpty(sourceData))
                {
                    workInProgress[rowNumber] = ProcessOneRuleWithGUI(sourceData, providerNameTitleRegex);
                }
            }

            // Run next rule on workInProgress & replace workInProgress.
            for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
            {
                if (cancel) { break; }

                // Modify & store as workInProgress.
                if (!string.IsNullOrEmpty(workInProgress[rowNumber]))
                {
                    workInProgress[rowNumber] = ProcessOneRuleWithGUI(workInProgress[rowNumber], drNameRegex);
                }
            }

            // Run final rule on workInProgress & save to target cell.
            for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
            {
                if (cancel) { break; }

                target = (Range)worksheet.Cells[rowNumber, aliasedColumn.Column];

                // Modify & stuff into target cell.
                if (!string.IsNullOrEmpty(workInProgress[rowNumber]))
                {
                    target.Value = ProcessOneRuleWithGUI(workInProgress[rowNumber], nameRegex);
                }
            }
        }

        internal void ObscureDateTime(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            // Instantiate random number generator and random day, time offsets.
            Random rnd = new Random();
            dayOffset = rnd.Next(-7, 7);
            monthOffset = rnd.Next(-2, 2);
            int hourOffset = rnd.Next(-3, 3);
            int minuteOffset = rnd.Next(-20, 20);
            deltaT = new TimeSpan(hourOffset, minuteOffset, 0);

            // Instantiate reusable Regexes.
            dateOnlyRegex = new Regex(dateOnlyPattern);
            dateTimeRegex = new Regex(dateTimePattern);
            monthDayOnlyRegex = new Regex(monthDayOnlyPattern);
            monthSpelledOutRegex = new Regex(monthSpelledOutPattern);

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                string newColumnName = selectedColumnName + " (Date/Time Altered)";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng, 
                                                                 newColumnName: newColumnName, 
                                                                 side: InsertSide.Right);

                string sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;

                    // Modify & stuff into target cell.
                    if (!string.IsNullOrEmpty(sourceData))
                    {
                        sourceData = ProcessOneRule(sourceData, dateTimeRegex, TweakDateTime);
                        sourceData = ProcessOneRule(sourceData, dateOnlyRegex, TweakDateOnly);
                        sourceData = ProcessOneRule(sourceData, monthDayOnlyRegex, TweakMonthDay);
                        sourceData = ProcessOneRule(sourceData, monthSpelledOutRegex, TweakMonthSpelledOut);
                    }

                    target.Value = sourceData;
                }
            }
        }

        private string ProcessOneRule(string sourceData, Regex regex, Func<string, string> convert)
        {
            Match match = regex.Match(sourceData);
            string targetData = string.Empty;

            while (match.Success)
            {
                string beforeMatch = sourceData.Substring(0, match.Index);
                targetData += beforeMatch;
                string matchedWord = match.Value.ToString();

                targetData += convert(matchedWord);

                // Trim to just what's AFTER the match.
                sourceData = sourceData.Substring(match.Index + match.Value.Length);
                match = regex.Match(sourceData);
            }

            // Append whatever's left over.
            targetData += sourceData;

            return targetData;
        }

        private string ProcessOneRuleWithGUI(string sourceData, Regex regex)
        {
            Match match = regex.Match(sourceData);
            string targetData = string.Empty;

            while (match.Success)
            {
                string matchedWord = match.Value.ToString().Trim().TrimEnd(',');

                string compiledPattern = wordsBeforePattern + matchedWord + wordsAfterPattern;
                Regex contextRegex = new Regex(compiledPattern);
                Match contextMatch = contextRegex.Match(sourceData);
                string leftContext = string.Empty;
                string rightContext = string.Empty;

                if (contextMatch.Success && contextMatch.Groups.Count > 2)
                {
                    leftContext = contextMatch.Groups["left"].Value.ToString();
                    rightContext = contextMatch.Groups["right"].Value.ToString();
                }

                SelectionResult selectionResult = EncodeProviderName(matchedWord, leftContext, rightContext);

                if (cancel) { break; }

                // Where's the detected word in the text?
                int index = sourceData.IndexOf(selectionResult.wordToReplace);
                string beforeText = sourceData.Substring(0, index);

                // Build the output.
                targetData += beforeText;

                // Apply the user's decision.
                if (selectionResult.replace)
                {
                    // Replace the word.
                    targetData += selectionResult.alias;
                }
                else
                {
                    // We're not replacing it.
                    targetData += selectionResult.wordToReplace;
                }

                // Trim to just what's AFTER the match.
                sourceData = sourceData.Substring(index + selectionResult.wordToReplace.Length);

                // Rerun the rule on rest of the data.
                match = regex.Match(sourceData);
            }

            // Append whatever's left over.
            targetData += sourceData;

            return targetData;
        }

        private string StringToHash(string sourceData)
        {
            string hashString = string.Empty;
     
            // Create a byte array from source data.
            byte[] tmpSource = ASCIIEncoding.ASCII.GetBytes(sourceData);

            // Initialize a SHA256 hash object.
            using (SHA256 mySHA256 = SHA256.Create())
            {
                byte[] tmpHash = mySHA256.ComputeHash(tmpSource);
                hashString = ByteArrayToString(tmpHash);
            }

            return hashString;
        }

        private string TweakDateOnly(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            string convertedDateString = payloadTweaked.ToString("M/d/yyyy");

            // Special case: did the Regex absorb a trailing period or space?
            if (dateString.EndsWith("."))
            {
                convertedDateString += ".";
            }

            if (dateString.EndsWith(" "))
            {
                convertedDateString += " ";
            }

            return convertedDateString;
        }

        private string TweakDateTime(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset) + deltaT;
            return payloadTweaked.ToString("M/d/yyyy h:mm tt");
        }

        private string TweakMonthDay(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            return payloadTweaked.ToString("M/d");
        }

        private string TweakMonthSpelledOut(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddMonths(monthOffset);
            return payloadTweaked.ToString("MMMM yyyy");
        }
    }
}