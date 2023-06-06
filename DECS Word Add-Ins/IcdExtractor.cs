using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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

        // Accommodate both condition - code ("Alzheimer’s disease – G30") and code = condition ("F12 = cannabis") formats.
        private readonly string[] LINE_PATTERNS = { @"(?<condition>[\w ',]+) +[-=:]? *(?<code>[A-Z]\d+[A-Z]?\.?\d*)",
                                                    @"(?<code>[A-Z]\d+[A-Z]?\.?\d*) +[-=:]? *(?<condition>[\w ',]+)" };

        // The numerical part of an ICD-10 code.
        private const string NUMBER_PATTERN = @"\d+\.?\d*";

        // Detect instruction like "M30 – M36".
        private const string SERIES_PATTERN = @"([A-Z]\d+\.?\d*) +[-=:] *([A-Z]\d+\.?\d*)";
        //
        // SQL snippets
        //
        private const string EXTRA_CODE_LINE = "\r\n                                          OR CODE LIKE ";
        private const string PREAMBLE = "\r\nSELECT DISTINCT\r\n        MRN,";
        private const string PREFIX = "\r\n        CASE\r\n              WHEN\r\n                    (\r\n                        SELECT TOP 1\r\n                               DX_ID\r\n                        FROM   problem_list\r\n                        WHERE  PAT_ID = pat.PAT_ID\r\n                        AND    DX_ID IN\r\n                               (\r\n                                      SELECT DX_ID\r\n                                      FROM   EDG_CURRENT_ICD10\r\n                                      WHERE  CODE LIKE ";
        private const string SUFFIX = "\r\n                    ) IS NOT NULL THEN 'Y'\r\n                      ELSE 'N'\r\n        END AS ";

        private Regex alphaRegex;
        private Regex codeRegex;
        private Regex[] lineRegexes;
        private Regex numberRegex;
        private Regex seriesRegex;

        internal IcdExtractor()
        {
            BuildRegex();
        }
        // Create all the reusable Regex objects.
        private void BuildRegex()
        {
            this.alphaRegex = new Regex(ALPHA_PATTERN);
            this.codeRegex = new Regex(CODE_PATTERN);

            this.lineRegexes = new Regex[LINE_PATTERNS.Length];

            for (int i = 0; i < LINE_PATTERNS.Length; i++)
            {
                this.lineRegexes[i] = new Regex(LINE_PATTERNS[i]);
            }

            this.numberRegex = new Regex(NUMBER_PATTERN);
            this.seriesRegex = new Regex(SERIES_PATTERN);
        }
        // Expand text like "M30 - M35" into a comma-separated string "M30, M31, M32, M33, M34, M35".
        private string ExpandSeries(string text)
        {
            string expanded_text = text;
            string alpha;
            int endNumber;
            int startNumber;

            MatchCollection matches = this.seriesRegex.Matches(text);

            foreach (Match match in matches)
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    // The thing we need to replace.
                    string series_definition = match.Groups[0].Value;
                    Match start_match = this.numberRegex.Match(match.Groups[1].Value);
                    
                    if (!Int32.TryParse(start_match.Groups[0].Value, out startNumber)) continue;

                    Match end_match = this.numberRegex.Match(match.Groups[2].Value);

                    if (!Int32.TryParse(end_match.Groups[0].Value, out endNumber)) continue;

                    int sequence_count = endNumber - startNumber;
                    Match alpha_match = this.alphaRegex.Match(match.Groups[1].Value);
                    alpha = alpha_match.Groups[0].Value;

                    int[] codeNumberSequence = Enumerable.Range(startNumber, sequence_count + 1).ToArray();
                    string[] codes_with_alpha = codeNumberSequence.Select(i => alpha + i.ToString()).ToArray();
                    string codes = String.Join(",", codes_with_alpha);
                    expanded_text = text.Replace(series_definition, codes);
                    break;
                }
            }

            return expanded_text;
        }
        // Handle a paragraph, which is probably just one line (since it ends with a newline.)
        internal void ProcessParagraph(string text, StreamWriter writer)
        {
            // Look for all the ICD codes in the paragraph (to be able to handle things like "M30, M31, M32").
            MatchCollection code_matches = this.codeRegex.Matches(text);

            if (code_matches.Count > 0)
            {
                bool found_match = false;

                foreach (Regex lineRegex in this.lineRegexes)
                {
                    Match line_match = lineRegex.Match(text);
                    string condition_name = "";

                    if (line_match.Success)
                    {
                        bool first_match = true;
                        condition_name = line_match.Groups["condition"].Value;

                        if (condition_name == null) continue;

                        if (Utilities.IsJustListOfCodes(name: condition_name, matches: code_matches)) continue;

                        foreach (Match code_match in code_matches)
                        {
                            string code_value = code_match.Groups[0].Value;

                            if (code_value == null) continue;

                            if (FALSE_CODES.Contains(code_value)) continue;

                            found_match = true;

                            if (first_match)
                            {
                                writer.Write(PREFIX + "'" + code_value + "%'");
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
                        writer.WriteLine(SUFFIX + Utilities.CleanNameForSql(condition_name));
                        break;
                    }
                }
            }
        }
        // Main method. Accepts a Document object & writes out the .sql file.
        internal void Scan(Document doc)
        {
            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(input_filename: doc.FullName, filetype: ".sql");

            writer.WriteLine(PREAMBLE);

            foreach (Paragraph para in doc.Paragraphs)
            {
                if (para == null) continue;

                string text = para.Range.Text.ToString().Trim();

                if (text == null) continue;

                text = Utilities.CleanText(text);

                string textExpanded = ExpandSeries(text);

                ProcessParagraph(textExpanded, writer);
            }
            
            writer.Close();
            Utilities.ShowResults(outputFilename);
        }
    }
}