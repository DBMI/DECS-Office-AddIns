using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Web;

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
