using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        internal int startOffset()
        {
            return startRowOffset;
        }

    }
}
