using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static log4net.Appender.RollingFileAppender;

namespace DECS_Excel_Add_Ins
{
    internal class DateRangeColumns
    {
        private string _name;
        private DateRange _dateRange;
        private Range _startDateColumn;
        private int _startDateColumnOffset;
        private Range _endDateColumn;
        private int _endDateColumnOffset;
        private Range _topLeftCorner;
        private Worksheet _worksheet;

        internal DateRangeColumns(Range startDateColumn, Range endDateColumn, string name)
        {
            _startDateColumn = startDateColumn;
            _endDateColumn = endDateColumn;
            _name = name;

            // Derived values.
            _worksheet = _startDateColumn.Worksheet;
            _topLeftCorner = (Range)_worksheet.Cells[1, 1];
            _startDateColumnOffset = startDateColumn.Column - 1;
            _endDateColumnOffset = endDateColumn.Column - 1;

            // Initialize DateRange object.
            _dateRange = GetDates();
        }

        internal bool CanMergeDates(int rowOffset)
        {
            DateRange newDateRange = GetDates(rowOffset);
            
            return _dateRange.Contiguous(newDateRange);
        }

        internal string EndColumnName()
        {
            return Convert.ToString(_topLeftCorner.Offset[0, _endDateColumnOffset].Value2);
        }

        internal DateTime EndDate()
        {
            return _dateRange.End();
        }

        private DateRange GetDates(int rowOffset = 1)
        {
            DateRange dateRange = null;

            // Pull the start & end dates from the ranges.
            string currentValue;
            DateTime? startDate = null;
            DateTime? endDate = DateTime.MaxValue;

            try
            {
                currentValue = _topLeftCorner.Offset[rowOffset, _startDateColumnOffset].Value.ToString();
                startDate = Utilities.ConvertExcelDate(currentValue);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            try
            {
                currentValue = _topLeftCorner.Offset[rowOffset, _endDateColumnOffset].Value.ToString();
                endDate = Utilities.ConvertExcelDate(currentValue);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            if (startDate.HasValue && endDate.HasValue)
            {   
                dateRange = new DateRange(startDate.Value, endDate.Value);
            }

            return dateRange;
        }

        internal string StartColumnName()
        {
            return Convert.ToString(_topLeftCorner.Offset[0, _startDateColumnOffset].Value2);
        }

        internal DateTime StartDate()
        {
            return _dateRange.Start();
        }

        internal void UpdateDates(int rowOffset)
        {
            DateRange newDateRange = GetDates(rowOffset);
            _dateRange = _dateRange.Merge(newDateRange);
        }
    }
}
