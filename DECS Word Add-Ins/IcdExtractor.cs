using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DecsWordAddIns
{
    internal class IcdExtractor
    {
        // Extract the alpha part ("M") of an ICD-10 code ("M30").
        private const string ALPHA_PATTERN = @"[A-Z]";

        // Recognize a full ICD-10 code (like "G47.33")
        private const string CODE_PATTERN = @"[A-Z]\d+[A-Z]?\.?\d*";

        // These strings may be present in the Statement of Work documents, but aren't ICD codes.
        private readonly string[] FALSE_CODES = { "A1C", "L1T", "R001442" };

        private readonly string LEADING_PAREN = @"^ ?\)? ?,? ?";

        // Accommodate formats:
        //  condition - code "Alzheimer’s disease – G30"
        //  condition: code "Unspecified sensorineural hearing loss( ICD-10-CM: H90.5 )"
        //  code = condition "F12 = cannabis"
        private readonly string[] LINE_PATTERNS =
        {
            @"(?<condition>[\w ',]+) +[-=:]? *(?<code>[A-Z]\d+[A-Z]?\.?\d*)",
            @"(?<condition>[\w, \-\(\)]+) *(?:\(CMS-HCC\))?\( *ICD-10-CM: *(?<code>[A-Z]\d+[A-Z\.\d\*]*)",
            @"(?<code>[A-Z]\d+[A-Z]?\.?\d*) +[-=:]? *(?<condition>[\w ',]+)"
        };

        // The numerical part of an ICD-10 code.
        private const string NUMBER_PATTERN = @"\d+\.?\d*";

        // Detect instruction like "M30 – M36".
        private const string SERIES_PATTERN = @"([A-Z]\d+\.?\d*) +[-=:] *([A-Z]\d+\.?\d*)";
        //
        // SQL snippets
        //
        private const string CASE_PREAMBLE = "\r\nSELECT DISTINCT\r\n        MRN,";
        private const string CASE_PREFIX =
            "\r\n        CASE\r\n              WHEN\r\n                    (\r\n                        SELECT TOP 1\r\n                               DX_ID\r\n                        FROM   problem_list\r\n                        WHERE  PAT_ID = pat.PAT_ID\r\n                        AND    DX_ID IN\r\n                               (\r\n                                      SELECT DX_ID\r\n                                      FROM   EDG_CURRENT_ICD10\r\n                                      WHERE  CODE LIKE ";
        private const string CASE_SUFFIX =
            "\r\n                    ) IS NOT NULL THEN 'Y'\r\n                      ELSE 'N'\r\n        END AS ";
        private const string EXTRA_CODE_LINE =
            "\r\n                                          OR CODE LIKE ";
        private const string LIKE_PREFIX = "OR CODE LIKE ";
        private const string LIST_PREAMBLE = "\r\n\r\nDROP TABLE IF EXISTS #ICD_CODES;\r\nSELECT DISTINCT DX_ID\r\nINTO #ICD_CODES\r\nFROM EDG_CURRENT_ICD10\r\nWHERE ";
        private const string LIST_PREFIX = "   CODE IN (";
        private Regex alphaRegex;
        private Regex codeRegex;
        private IcdStyle icdStyle = IcdStyle.Case;
        private Regex[] lineRegexes;
        private Regex numberRegex;
        private Regex seriesRegex;

        internal IcdExtractor()
        {
            BuildRegex();
            
            // Ask user for output style & return quietly if they cancel.
            if (!AskStyle())
            {
                return;
            }
        }

        private bool AskStyle()
        {
            bool success = false;

            // Ask user whether they want output in CASE style or LIST style.
            using (var form = new StyleSelectionForm())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    icdStyle = form.icdStyle;
                    success = true;
                }
            }

            return success;
        }

        // Create all the reusable Regex objects.
        private void BuildRegex()
        {
            alphaRegex = new Regex(ALPHA_PATTERN);
            codeRegex = new Regex(CODE_PATTERN);

            lineRegexes = new Regex[LINE_PATTERNS.Length];

            for (int i = 0; i < LINE_PATTERNS.Length; i++)
            {
                lineRegexes[i] = new Regex(LINE_PATTERNS[i]);
            }

            numberRegex = new Regex(NUMBER_PATTERN);
            seriesRegex = new Regex(SERIES_PATTERN);
        }

        // Expand text like "M30 - M35" into a comma-separated string "M30, M31, M32, M33, M34, M35".
        private string ExpandSeries(string text)
        {
            string expanded_text = text;
            string alpha;
            int endNumber;
            int startNumber;

            MatchCollection matches = seriesRegex.Matches(text);

            foreach (Match match in matches)
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    // The thing we need to replace.
                    string series_definition = match.Groups[0].Value;
                    Match start_match = numberRegex.Match(match.Groups[1].Value);

                    if (!int.TryParse(start_match.Groups[0].Value, out startNumber))
                        continue;

                    Match end_match = numberRegex.Match(match.Groups[2].Value);

                    if (!int.TryParse(end_match.Groups[0].Value, out endNumber))
                        continue;

                    int sequence_count = endNumber - startNumber;
                    Match alpha_match = alphaRegex.Match(match.Groups[1].Value);
                    alpha = alpha_match.Groups[0].Value;

                    int[] codeNumberSequence = Enumerable
                        .Range(startNumber, sequence_count + 1)
                        .ToArray();
                    string[] codes_with_alpha = codeNumberSequence
                        .Select(i => alpha + i.ToString())
                        .ToArray();
                    string codes = string.Join(",", codes_with_alpha);
                    expanded_text = text.Replace(series_definition, codes);
                    break;
                }
            }

            return expanded_text;
        }

        // Handle a paragraph, which is probably just one line (since it ends with a newline.)
        private void ProcessParagraphCase(string text, StreamWriter writer)
        {
            if (text == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            // Only once we know there is text to process.
            string textExpanded = ExpandSeries(text);

            // Look for all the ICD codes in the paragraph (to be able to handle things like "M30, M31, M32").
            MatchCollection code_matches = codeRegex.Matches(textExpanded);

            if (code_matches.Count > 0)
            {
                bool found_match = false;

                foreach (Regex lineRegex in lineRegexes)
                {
                    Match line_match = lineRegex.Match(text);
                    string condition_name = "";

                    if (line_match.Success)
                    {
                        bool first_match = true;
                        condition_name = line_match.Groups["condition"].Value;

                        if (condition_name == null)
                            continue;

                        if (
                            Utilities.IsJustListOfCodes(name: condition_name, matches: code_matches)
                        )
                            continue;

                        foreach (Match code_match in code_matches)
                        {
                            string code_value = code_match.Groups[0].Value;

                            if (code_value == null)
                                continue;

                            if (FALSE_CODES.Contains(code_value))
                                continue;

                            found_match = true;

                            if (first_match)
                            {
                                writer.Write(CASE_PREFIX + "'" + code_value + "%'");
                                first_match = false;
                            }
                            else
                            {
                                writer.Write(EXTRA_CODE_LINE + "'" + code_value + "%'");
                            }

                            // Remove code values ("J42") from the code name.
                            condition_name = condition_name.Replace(code_value, "");
                            condition_name = condition_name.Replace(",", "");
                            condition_name = condition_name.Trim();
                        }
                    }

                    if (found_match)
                    {
                        writer.WriteLine(") -- " + condition_name);
                        writer.WriteLine(CASE_SUFFIX + Utilities.CleanNameForSql(condition_name));
                        break;
                    }
                }
            }
        }

        private void ProcessParagraphList(string text, StreamWriter writer)
        {
            if (text == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            // Only once we know there is text to process.
            writer.WriteLine(LIST_PREAMBLE);

            List<string> likeCodes = new List<string>();
            List<string> likeConditions = new List<string>();

            foreach (Regex lineRegex in lineRegexes)
            {
                List<Match> lineMatches = lineRegex.Matches(text).Cast<Match>().Where(m => m.Success).ToList();
                
                if (lineMatches.Count == 0)
                {
                    continue;
                }

                for (int i = 0; i < lineMatches.Count; i++)
                {
                    Match thisMatch = lineMatches[i];

                    if (thisMatch.Success)
                    {
                        string codeValue = thisMatch.Groups["code"].Value;
                        string conditionName = thisMatch.Groups["condition"].Value;
                        
                        // Strip off leftovers from previous condition name.
                        conditionName = Regex.Replace(conditionName, LEADING_PAREN, "");

                        if (conditionName == null || codeValue == null)
                            continue;

                        // As we find codes like "G20.*", save them
                        // & we'll process them all at once at the end.
                        if (codeValue.Contains("*"))
                        {
                            likeCodes.Add(codeValue);
                            likeConditions.Add(conditionName);
                            continue;
                        }

                        if (i == 0)
                        {
                            if (i < (lineMatches.Count - 1))
                            {
                                writer.Write(LIST_PREFIX + "'" + codeValue + "',  --- " + conditionName);
                            }
                            else
                            {
                                writer.Write(LIST_PREFIX + "'" + codeValue + "'  --- " + conditionName);
                            }
                        }
                        else
                        {
                            if (i < (lineMatches.Count - 1))
                            {
                                writer.Write("\r\n            '" + codeValue + "',  --- " + conditionName);
                            }
                            else
                            {
                                writer.Write("\r\n            '" + codeValue + "')  --- " + conditionName);
                            }
                        }
                    }
                }
            }

            if (likeCodes.Count > 0)
            {
                // Now list all the "Like" rules.
                for (int i = 0; i < likeCodes.Count - 1; i++)
                {
                    string codeValue = likeCodes[i].ToString();
                    string conditionName = likeConditions[i].ToString();
                    writer.Write(")\r\n" + LIKE_PREFIX + "'" + codeValue.Replace("*", "%") + "'   --- " + conditionName);
                }

                // Different formatting for last one.
                string lastCodeValue = likeCodes[likeCodes.Count - 1].ToString();
                string lastConditionName = likeConditions[likeCodes.Count - 1].ToString();
                writer.Write(")\r\n" + LIKE_PREFIX + "'" + lastCodeValue.Replace("*", "%") + "';   --- " + lastConditionName);
            }
        }

        // Main method. Accepts a Document object & writes out the .sql file.
        internal void Scan(Document doc)
        {
            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                input_filename: doc.FullName,
                filetype: ".sql"
            );

            if (icdStyle == IcdStyle.Case)
            {
                writer.WriteLine(CASE_PREAMBLE);
            }

            // Get just the selected region or (if no selection), the entire document.
            List<string> textBlocks = Utilities.SelectedText(doc);

            if (textBlocks == null || textBlocks.Count == 0)
            {
                return;
            }            

            foreach (string textBlock in textBlocks)
            {
                if (string.IsNullOrEmpty(textBlock))
                {
                    continue;
                }

                // Remove junk that confuses the Regular Expressions.
                string textBlockCleaned = Utilities.CleanText(textBlock);

                switch (icdStyle)
                {
                    case IcdStyle.Case:
                        ProcessParagraphCase(textBlockCleaned, writer);
                        break;

                    case IcdStyle.List:
                        ProcessParagraphList(textBlockCleaned, writer);
                        break;
                }
            }

            writer.Close();
            Process.Start(outputFilename);
        }
    }
}
