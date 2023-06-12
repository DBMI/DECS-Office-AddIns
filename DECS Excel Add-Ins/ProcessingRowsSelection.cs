using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    internal class ProcessingRowsSelection
    {
        private bool allRows;
        private Excel.Range rows; // From (Excel.Range)this.application.Selection.Rows;
        private string reason;

        public ProcessingRowsSelection(Excel.Range rows, string reason, bool allRows = false)
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

        internal Excel.Range GetRows()
        {
            return rows;
        }

        internal int NumRows()
        {
            return rows.Count;
        }

        public override string ToString()
        {
            return this.rows.Count.ToString() + " rows selected.";
        }
    }
}
