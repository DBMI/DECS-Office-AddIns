using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class DateRange
    {
        private DateTime _start;
        private DateTime _end;
        private const int oneDay = 1;

        internal DateRange(DateTime start, DateTime end)
        {
            // Don't allow reversed dates (start > end).
            if (start <= end)
            {
                _start = start;
                _end = end;
            }
            else
            {
                // Automatically swap if they're reversed.
                _start = end;
                _end = start;
            }
        }

        /// <summary>
        /// Creates a DateRange object from string like '01/05-01/11'.
        /// </summary>
        /// <param name="dateContent">string</param>
        internal DateRange(string dateContent, int assumedYear)
        {
            Regex regex = new Regex(@"(?<month>\d{1,2})\/(?<day>\d{1,2})");
            string[] dateParts = dateContent.Split('-');

            if (dateParts.Length == 2) 
            {
                Match start_match = regex.Match(dateParts[0]);

                if (start_match.Success)
                {
                    if (int.TryParse(start_match.Groups["day"].Value, out int day) &&
                        int.TryParse(start_match.Groups["month"].Value, out int month))
                    {
                        _start = new DateTime(assumedYear, month, day);
                    }                    
                }

                Match end_match = regex.Match(dateParts[1]);

                if (end_match.Success)
                {
                    if (int.TryParse(end_match.Groups["day"].Value, out int day) &&
                        int.TryParse(end_match.Groups["month"].Value, out int month))
                    {
                        _end = new DateTime(assumedYear, month, day);
                    }
                }

                // Special handling for end of the year like: "12/29-01/04"
                if (_start > _end)
                {
                    _start = _start.AddYears(-1);   // Move it to previous year.
                }
            }
        }

        /// <summary>
        /// Bump the year by one 
        /// (for when we discover after the fact that we used the wrong value for assumedYear.)
        /// </summary>

        internal void AddYear()
        {
            _start = _start.AddYears(1);
            _end = _end.AddYears(1);
        }

        /// <summary>
        /// Is THIS object entirely AFTER this other DateRange object?
        /// We'll tolerate differences of one day.
        /// So we'll consider that a DateRange that starts on 01 June
        /// is NOT "after" a DateRange that ends on 31 May.
        /// </summary>
        /// <param name="newDateRange">Another DateRange object.</param>

        private bool After(DateRange newDateRange)
        {
            if (newDateRange == null)
            {
                return false;
            }

            return (this.Start() - newDateRange.End()).Days > oneDay;
        }

        /// <summary>
        /// Is THIS object entirely BEFORE this other DateRange object?
        /// We'll tolerate differences of one day.
        /// So we'll consider that a DateRange that ends on 31 May
        /// is NOT "before" a DateRange that starts on 01 June.
        /// </summary>
        /// <param name="newDateRange">Another DateRange object.</param>

        private bool Before(DateRange newDateRange)
        {
            if (newDateRange == null)
            {
                return false;
            }

            return (newDateRange.Start() - this.End()).Days > oneDay;
        }

        /// <summary>
        /// Can this DateRange be combined with another DateRange object?
        /// </summary>
        /// <param name="newDateRange">Another DateRange object.</param>
        /// <param name="spanGaps">If there's a gap in date coverage, do we merge the objects? (default: false)</param>
        
        internal bool Contiguous(DateRange newDateRange, bool spanGaps = false)
        {
            if (spanGaps)
            {
                return true;
            }

            // Test for disjoint date ranges,
            // in which this object is entirely BEFORE or AFTER newDateRange object.
            return !(this.After(newDateRange)) && !(this.Before(newDateRange));
        }

        internal DateTime End()
        {
            return _end;
        }

        internal bool Valid()
        {
            return _start != null && _end != null;
        }

        /// <summary>
        /// Combine with another DateRange object.
        /// </summary>
        /// <param name="newDateRange">Another DateRange object.</param>
        
        internal DateRange Merge(DateRange newDateRange)
        {

            DateTime earlierStart;
            DateTime laterEnd;

            if (_start <= newDateRange.Start())
            {
                earlierStart = _start;
            }
            else
            {
                earlierStart = newDateRange.Start();
            }

            if (_end >= newDateRange.End())
            {
                laterEnd = _end;
            }
            else
            {
                laterEnd = newDateRange.End();
            }

            return new DateRange(earlierStart, laterEnd);
        }

        public static bool operator >(DateRange r1, DateRange r2)
        {
            if ((object)r1 == null)
            {
                return false;
            }

            if ((object)r2 == null)
            {
                return true;
            }

            return r1.Start() > r2.End();
        }

        public static bool operator <(DateRange r1, DateRange r2)
        {
            if ((object)r1 == null)
            {
                if ((object)r2 == null)
                {
                    return false;
                }

                return true;
            }

            if ((object)r2 == null)
            {
                return false;
            }

            return r1.End() < r2.Start();
        }

        internal DateTime Start()
        {
            return _start;
        }
    }
}
