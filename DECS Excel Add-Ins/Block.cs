using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    // Represents a series of rows in a sheet.
    internal class Block
    {
        // Offset from row 1.
        private int startRowOffset;
        private int endRowOffset;

        internal Block(int startRowOffset, int endRowOffset)
        {
            this.startRowOffset = startRowOffset;
            this.endRowOffset = endRowOffset;
        }

        internal int endOffset()
        {
            return endRowOffset;
        }

        private int numRows()
        {
            return endRowOffset - startRowOffset + 1;
        }

        internal string rowList()
        {
            // Convert from offsets to row numbers.
            return (startRowOffset + 1).ToString() + ":" + (endRowOffset + 1).ToString();
        }

        // Create a new Block somewhere else the same size as this one.
        internal Block sameSize(int newStartingRowOffset)
        {
            return new Block(newStartingRowOffset, newStartingRowOffset + numRows() - 1);
        }

        internal void shade(Worksheet worksheet, XlRgbColor shade)
        {
            int lastColInSheet = worksheet.UsedRange.Columns.Count;

            Range startCell = (Range)worksheet.Cells[startRowOffset + 1, 1];
            Range endCell = (Range)worksheet.Cells[endRowOffset + 1, lastColInSheet];
            Range theseRows = (Range)worksheet.Range[startCell, endCell];
            theseRows.Interior.Color = shade;
        }

        internal int startOffset()
        {
            return startRowOffset;
        }

    }
}
