using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Priority of triage action.
    */
    // https://stackoverflow.com/a/479417/18749636
    internal enum TriagePriority
    {
        [Description("0-2 weeks")]
        High,

        [Description("2-4 weeks")]
        Medium,

        [Description("4+ weeks")]
        Routine,

        [Description("Unknown")]
        Unknown
    }

    internal class TimeSorter
    {
        private Microsoft.Office.Interop.Excel.Application application;
        Dictionary<string, Range> columnNamesDict;
        private int lastRow;
        private Dictionary<string, string> monthAbbreviations;
        private Dictionary<string, string> seasons;
        private Dictionary<string, int> weeks;
        private Range selectedDateColumnRng;
        private Range selectedTimeTextColumnRng;
        private FollowUpTimeframeThresholds thresholds;

        internal TimeSorter()
        {
            application = Globals.ThisAddIn.Application;
            BuildMonthDictionary();
            BuildSeasonsDictionary();
            BuildWeeksDictionary();

            TimeSorterSettings settings = new TimeSorterSettings();
            thresholds = settings.Thresholds();
        }

        private void BuildMonthDictionary()
        {
            monthAbbreviations = new Dictionary<string, string>();

            monthAbbreviations.Add("Jan", "January");
            monthAbbreviations.Add("Feb", "February");
            monthAbbreviations.Add("Mar", "March");
            monthAbbreviations.Add("Apr", "April");
            monthAbbreviations.Add("Jun", "June");
            monthAbbreviations.Add("Jul", "July");
            monthAbbreviations.Add("Aug", "August");
            monthAbbreviations.Add("Sept", "September");
            monthAbbreviations.Add("Oct", "October");
            monthAbbreviations.Add("Nov", "November");
            monthAbbreviations.Add("Dec", "December");
        }

        private void BuildSeasonsDictionary()
        {
            seasons = new Dictionary<string, string>();

            // What's the end of the season?
            seasons.Add("spring", "May");
            seasons.Add("summer", "August");
            seasons.Add("autumn", "November");
            seasons.Add("fall", "November");
            seasons.Add("winter", "February");
        }

        private void BuildWeeksDictionary()
        {
            weeks = new Dictionary<string, int>();

            // What's a date in this week?
            weeks.Add("beginning", 15);
            weeks.Add("early", 15);
            weeks.Add("first", 5);
            weeks.Add("second", 12);
            weeks.Add("third", 19);
            weeks.Add("fourth", 26);
            weeks.Add("fifth", 30);
            weeks.Add("end", 30);
        }

        private bool FindSelectedDateColumn(Worksheet worksheet)
        {
            bool success = false;

            // What's the column we DON'T want? The time text column.
            string timeTextColumnName = Utilities.FindColumnName(columnNamesDict, selectedTimeTextColumnRng);

            // Ask user to select one column.
            List<string> columnNames = Utilities.GetColumnNames(worksheet);
            columnNames.Remove(timeTextColumnName);

            using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames,
                                                                    headline: "Choose Date Column",
                                                                    MultiSelect: false))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    selectedDateColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                    success = true;
                }
                else if (result == DialogResult.Cancel)
                {
                    // Then we're done here.
                    return success;
                }
            }

            return success;
        }

        private bool FindSelectedTimeTextColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedTimeTextColumnRng = Utilities.GetSelectedCol(application);

            if (selectedTimeTextColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames,
                                                                        headline: "Choose Time Text Column",
                                                                        MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        selectedTimeTextColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
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

        private DateTime? NextMonth(string month, DateTime refDate, int day = 1)
        {
            DateTime? dateTranslated = null;

            // When did the user write the time text?
            int year = refDate.Year;

            // Form date string (like "July 20 2025") from the pieces,
            // assuming the month/day "July 20" refers to the same year
            // in which the user wrote the note.
            string dateString = month + " " + day.ToString() + " " + year.ToString();

            if (DateTime.TryParse(dateString, out DateTime dateTimeValue))
            {
                dateTranslated = dateTimeValue;
            }
            else
            {
                return null;
            }

            // Maybe when the user (in July 2025) wrote
            // about "Jan 14th" they meant the *following* January?
            while (dateTranslated < refDate)
            {
                year += 1;
                dateString = month + " " + day.ToString() + " " + year.ToString();

                if (DateTime.TryParse(dateString, out DateTime dateTimeValueNextYear))
                {
                    dateTranslated = dateTimeValueNextYear;
                }
                else
                {
                    return null;
                }
            }

            return dateTranslated;
        }

        private DateTime? ParseMonth(string timeText, DateTime noteWrittenDate)
        {
            DateTime? deadlineDate = null;

            // Strip off preamble like in 'mid'March.
            Regex regex = new Regex(@"(?:early|late|mid)?\s*(?<month>\w{3,})\s*(?<year>\d{4})?");
            Match match = regex.Match(timeText);

            if (match.Success)
            {
                // Maybe it's an abbreviated month?
                string month = TranslateMonthAbbreviation(match.Groups["month"].Value);

                // Is there year info?
                if (int.TryParse(match.Groups["year"].Value, out int year))
                {
                    // Form month year string (like "July 2024") from the pieces.
                    string monthYearString = month + " " + year.ToString();

                    if (DateTime.TryParse(monthYearString, out DateTime dateTimeValue))
                    {
                        deadlineDate = dateTimeValue;
                    }
                }
                else
                {
                    // See if the month belongs to this year or the next.
                    deadlineDate = NextMonth(month, noteWrittenDate);
                }
            }

            return deadlineDate;
        }

        private DateTime? ParseMonthDay(string timeText, DateTime noteWrittenDate)
        {
            DateTime? deadlineDate = null;

            Regex regex = new Regex(@"(?<month>\w{3,})\s*(?<day_first>\d{1,2})(?:th)?\s*(?:or)?\s*(?<day_last>\d{1,2})?");
            Match match = regex.Match(timeText);

            if (match.Success && match.Groups.Count > 2)
            {
                int? day = null;

                if (int.TryParse(match.Groups["day_last"].Value, out int day_last))
                {
                    day = day_last;
                }
                else if (int.TryParse(match.Groups["day_first"].Value, out int day_first))
                {
                    day = day_first;
                }

                if (day.HasValue)
                {
                    // Maybe it's an abbreviated month?
                    string month = TranslateMonthAbbreviation(match.Groups["month"].Value);

                    // See if the month belongs to this year or the next.
                    deadlineDate = NextMonth(month, noteWrittenDate, day.Value);
                }
            }

            return deadlineDate;
        }

        private TriagePriority ParsePriority(string timeText)
        {
            TriagePriority priority = TriagePriority.Unknown;

            // Looking for text like "4-10 weeks".
            Regex regex = new Regex(@"(?<qty>\d{1,})\s+(?<units>\w{4,})");
            Match match = regex.Match(timeText);

            if (match.Success && match.Groups.Count > 2)
            {
                string units = match.Groups["units"].Value;

                if (int.TryParse(match.Groups["qty"].Value, out int quantity))
                {
                    switch (units.ToLower())
                    {
                        case "month":
                        case "months":
                        case "mon":
                        case "mons":

                            priority = TriagePriority.Routine;
                            break;

                        case "week":
                        case "weeks":
                        case "wk":
                        case "wks":

                            priority = thresholds.ParsePriority(quantity);
                            break;

                        default:

                            priority = TriagePriority.Unknown;
                            break;
                    }
                }
            }

            return priority;
        }

        // When the text says "routine" directly.
        private TriagePriority ParsePriorityDirect(string timeText)
        {
            TriagePriority priority = TriagePriority.Unknown;

            switch (timeText.ToLower())
            {
                case "routine":
                    priority = TriagePriority.Routine;
                    break;

                default:
                    priority = TriagePriority.Unknown;
                    break;
            }

            return priority;
        }

        private DateTime? ParseSeasonYear(string timeText)
        {
            DateTime? deadlineDate = null;

            Regex regex = new Regex(@"(?<season>\w{3,})\s+(?<year>\d{4})");
            Match match = regex.Match(timeText);

            if (match.Success && match.Groups.Count > 2)
            {
                string season = match.Groups["season"].Value.ToLower();

                try
                {
                    string month = seasons[season];

                    if (int.TryParse(match.Groups["year"].Value, out int year))
                    {
                        // Form month year string (like "July 2024") from the pieces.
                        string monthYearString = month + " " + year.ToString();

                        if (DateTime.TryParse(monthYearString, out DateTime dateTimeValue))
                        {
                            deadlineDate = dateTimeValue;
                        }
                    }
                }
                catch (System.Collections.Generic.KeyNotFoundException)
                {
                }
            }

            return deadlineDate;
        }


        private DateTime? ParseNumericWeek(string timeText, DateTime noteWrittenDate)
        {
            DateTime? deadlineDate = null;

            // Like "2nd week of July"
            Regex regex = new Regex(@"\s*(?<weekNumber>\d{1})(?:nd|rd|st|th)?\s*week\s*(?:of)?\s*(?<month>\w{3,})");
            Match match = regex.Match(timeText);

            if (match.Success && match.Groups.Count > 2)
            {
                if (int.TryParse(match.Groups["weekNumber"].Value, out int weekNumber))
                {
                    // Maybe it's an abbreviated month?
                    string month = TranslateMonthAbbreviation(match.Groups["month"].Value);

                    // Assume the end of the week was meant.
                    int day = weekNumber * 7;

                    // See if the month belongs to this year or the next.
                    deadlineDate = NextMonth(month, noteWrittenDate, day);
                }
            }

            return deadlineDate;
        }

        private DateTime? ParseTextWeek(string timeText, DateTime noteWrittenDate)
        {
            DateTime? deadlineDate = null;

            // Like "first week of July"
            Regex regex = new Regex(@"\s*(?<weekWord>\w+)\s*(?:week\s*)?(?:of)?\s*(?<month>\w{3,})");
            Match match = regex.Match(timeText);

            if (match.Success && match.Groups.Count > 2)
            {
                string weekWord = match.Groups["weekWord"].Value;
                int? day = null;

                try
                {
                    day = weeks[weekWord];
                }
                catch (System.Collections.Generic.KeyNotFoundException)
                {
                    return null;
                }

                if (day.HasValue)
                {
                    // Maybe it's an abbreviated month?
                    string month = TranslateMonthAbbreviation(match.Groups["month"].Value);

                    // See if the month belongs to this year or the next.
                    deadlineDate = NextMonth(month, noteWrittenDate, day.Value);
                }
            }

            return deadlineDate;
        }

        internal void Scan(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;
            string dateText;
            string timeText;
            TriagePriority priority;

            columnNamesDict = Utilities.GetColumnRangeDictionary(worksheet);

            if (FindSelectedTimeTextColumn(worksheet) && FindSelectedDateColumn(worksheet))
            {
                // Create column for priority value.
                Range priorityRng = Utilities.InsertNewColumn(selectedTimeTextColumnRng, "Priority");

                // Run down the time text column, parsing the text.
                for (int rowOffset = 1; rowOffset < lastRow; rowOffset++)
                {
                    priority = TriagePriority.Unknown;

                    // Text that says something like "within 8 weeks"
                    timeText = selectedTimeTextColumnRng.Offset[rowOffset].Value.ToString();

                    // When the triage note was written.
                    dateText = selectedDateColumnRng.Offset[rowOffset].Value.ToString();

                    if (DateTime.TryParse(dateText, out DateTime noteWrittenDate))
                    {
                        // Is there a date in the time text? ("by April 2025")
                        if (DateTime.TryParse(timeText, out DateTime deadlineDate))
                        {
                            TimeSpan delta = deadlineDate - noteWrittenDate;
                            priority = thresholds.ParsePriority(delta);
                        }
                        else
                        {
                            priority = ParsePriority(timeText);
                        }
                        //
                        // Were we unable to parse the text?
                        //
                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it just say "routine"?
                            priority = ParsePriorityDirect(timeText);
                        }

                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it contain a date like "mid to late July 2024"?
                            DateTime? possibleDate = ParseMonth(timeText, noteWrittenDate);

                            if (possibleDate != null)
                            {
                                TimeSpan delta = possibleDate.Value - noteWrittenDate;
                                priority = thresholds.ParsePriority(delta);
                            }
                        }

                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it contain a date like "July 10th"?
                            DateTime? possibleDate = ParseMonthDay(timeText, noteWrittenDate);

                            if (possibleDate != null)
                            {
                                TimeSpan delta = possibleDate.Value - noteWrittenDate;
                                priority = thresholds.ParsePriority(delta);
                            }
                        }

                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it contain a season like "Fall 2025"?
                            DateTime? possibleDate = ParseSeasonYear(timeText);

                            if (possibleDate != null)
                            {
                                TimeSpan delta = possibleDate.Value - noteWrittenDate;
                                priority = thresholds.ParsePriority(delta);
                            }
                        }

                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it contain text like "3rd week of July"?
                            DateTime? possibleDate = ParseNumericWeek(timeText, noteWrittenDate);

                            if (possibleDate != null)
                            {
                                TimeSpan delta = possibleDate.Value - noteWrittenDate;
                                priority = thresholds.ParsePriority(delta);
                            }
                        }

                        if (priority == TriagePriority.Unknown)
                        {
                            // Does it contain text like "third week of July"?
                            DateTime? possibleDate = ParseTextWeek(timeText, noteWrittenDate);

                            if (possibleDate != null)
                            {
                                TimeSpan delta = possibleDate.Value - noteWrittenDate;
                                priority = thresholds.ParsePriority(delta);
                            }
                        }

                        priorityRng.Offset[rowOffset].Value = priority.GetDescription();
                    }
                }
            }
        }

        private string TranslateMonthAbbreviation(string mon)
        {
            string month = mon;

            if (monthAbbreviations.ContainsKey(mon))
            {
                month = monthAbbreviations[mon];
            }

            return month;
        }
    }
}
