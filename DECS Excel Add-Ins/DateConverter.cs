using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class DateConverter
    {
        private IDictionary<string, string> supportedDateFormats;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal DateConverter()
        {
            supportedDateFormats = new Dictionary<string, string>();
            supportedDateFormats.Add("MM/dd/yyyy", "(\\d{1,2}\\/\\d{1,2}\\/\\d{4})");
            supportedDateFormats.Add("MM-dd-yyyy", "(\\d{1,2}-\\d{1,2}-\\d{4})");
            supportedDateFormats.Add("dd MMMM yyyy", "(\\d{1,2} \\w{3,9}\\.? \\d{4})");
            supportedDateFormats.Add("MMMM dd yyyy", "(\\w{3,9}\\.? \\d{1,2},? \\d{4})");
            supportedDateFormats.Add("MMMM dd", "(\\w{3,9}\\.? \\d{1,2})");
        }

        internal string Convert(string note, string desiredFormat)
        {
            foreach (KeyValuePair<string, string> entry in supportedDateFormats)
            {
                if (entry.Key == desiredFormat)
                    continue;

                foreach (Match match in Regex.Matches(note, entry.Value))
                {
                    if (match.Success)
                    {
                        string dateString = match.Value.ToString();
                        log.Debug("Rule matched: " + dateString);

                        if (DateTime.TryParse(dateString, out DateTime dateValue))
                        {
                            string dateConverted = dateValue.ToString(desiredFormat);
                            log.Debug("Converted '" + dateString + "' to '" + dateConverted + "'.");
                            note = note.Replace(dateString, dateConverted);
                        }
                    }
                }
            }

            return note;
        }

        internal List<string> SupportedDateFormats()
        {
            return new List<string>(supportedDateFormats.Keys);
        }
    }
}
