using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Holds information from the user's selection of which rows to process:
     * - bool @c allRows Are we doing all rows? (Or just a selection?)
     * - Range @c rows Range including all the rows to process.
     * - string @c reason Code-generated explanation like "Selected row outside data area."
     */
    internal class ProcessingRowsSelection
    {
        private bool allRows;
        private Excel.Range rows; // From (Excel.Range)application.Selection.Rows;
        private string reason;

        public ProcessingRowsSelection(Excel.Range _rows, string _reason, bool _allRows = false)
        {
            rows = _rows;
            reason = _reason;
            allRows = _allRows;
        }

        /// <summary>
        /// Allows external code to ask if we're processing all rows.
        /// </summary>
        /// <returns>bool</returns>
        internal bool AllRows()
        {
            return allRows;
        }

        /// <summary>
        /// Allows external code to ask reason for processing decision.
        /// </summary>
        /// <returns>bool</returns>
        internal string GetReason()
        {
            return reason;
        }

        /// <summary>
        /// Allows external code to get the Range of rows to process.
        /// </summary>
        /// <returns>bool</returns>
        internal Excel.Range GetRows()
        {
            return rows;
        }

        /// <summary>
        /// Allows external code to ask the number of rows to process.
        /// </summary>
        /// <returns>bool</returns>
        internal int NumRows()
        {
            return rows.Count;
        }

        //public override string ToString()
        //{
        //    return rows.Count.ToString() + " rows selected.";
        //}
    }
}
