using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace DecsWordAddIns
{
    internal class MsgFormatter
    {
        private readonly string[] BREAK_AFTER_PATTERNS = { @"RE:\s*\S+\s*", @"sent at \d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2} [AP]M (?:\w{3} )?-----\s*", @"(Subject:\s*(?!RE:)[\d\w\s\>\<\(\)-]{1,40}(?<!Dr)\.\s+)" };
        private readonly string[] BREAK_INSIDE_PATTERNS = { @"(Subject:\s*[\d\w\s\>\<\(\)\.-]{1,40})(\s(?:Dear|Good|Hello|Hi|HI|I|My)[\s,]+)" };
        private readonly string[] BREAK_BEFORE_PATTERNS = { @"""?----- Message", "From:", "Sent:", "To:", @"Subject:\s*", @"\d{11},", @"sent at \d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2} [AP]M (?:\w{3} )?-----\s*" };
        private readonly string[] SPACE_AFTER_PATTERNS = { @"(From:)(\S+)", @"(Sent:)(\S+)", @"(To:)(\S+)", @"(Subject:)(\S+)", @"(RE:)(\S+)", @"(\w)\.(\w)" };

        public MsgFormatter()
        {
        }

        /// <summary>
        /// Main method: Formats the document to separate messages for human reviewers.
        /// </summary>
        /// <param name="doc">Word @c Document object</param>
        public void Format(Document doc)
        {
            string allText = doc.Range().Text;

            // Remove excess whitespace.
            allText = Regex.Replace(allText, @"(\S) {2,}", "$1 ");

            foreach (string breakAfterPattern in BREAK_AFTER_PATTERNS)
            {
                allText = Regex.Replace(allText, breakAfterPattern, "$&" + Environment.NewLine);
            }

            // https://stackoverflow.com/a/38168829/18749636
            foreach (string breakBeforePattern in BREAK_BEFORE_PATTERNS)
            {
                allText = Regex.Replace(allText, breakBeforePattern, Environment.NewLine + "$&");
            }

            foreach (string breakInsidePattern in BREAK_INSIDE_PATTERNS)
            {
                allText = Regex.Replace(allText, breakInsidePattern, "$1" + Environment.NewLine + "$2");
            }

            // https://stackoverflow.com/a/38168829/18749636
            foreach (string spaceAfterPattern in SPACE_AFTER_PATTERNS)
            {
                allText = Regex.Replace(allText, spaceAfterPattern, "$1 $2");
            }

            doc.Range().Text = allText;
        }
    }
}