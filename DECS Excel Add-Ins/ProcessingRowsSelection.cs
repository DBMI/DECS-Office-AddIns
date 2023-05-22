using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    internal class ProcessingRowsSelection
    {
        private bool allRows;
        private List<int> rows;
        private string reason;

        public ProcessingRowsSelection(List<int> rows, string reason, bool allRows = false)
        {
            this.rows = rows;
            this.reason = reason;
            this.allRows = allRows;
        }
        internal bool AllRows() 
        { 
            return allRows;
        }
        internal string GetReason()
        {
            return reason;
        }
        internal List<int> GetRows()
        {
            return rows;
        }
    }
}