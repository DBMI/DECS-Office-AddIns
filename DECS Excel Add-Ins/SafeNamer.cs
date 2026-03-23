
namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class to name a new worksheet that accounts for:
     *  1. 31-character limit
     *  2. Preventing collisions with existing names
     */
    internal class SafeNamer
    {
        private int nameIncrement = 1;
        private Microsoft.Office.Interop.Excel.Worksheet _worksheet;

        // Cut it a little short in case we need to append 1, 2, 3, etc.
        private const int MAX_LENGTH = 30;

        internal SafeNamer(Microsoft.Office.Interop.Excel.Worksheet worksheet) 
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Tries to give the worksheet the name you propose; increments the name if already taken.
        /// </summary>
        /// <param name="proposedName">string</param>
        /// <returns>_worksheet</returns>
        internal Microsoft.Office.Interop.Excel.Worksheet AssignName(string proposedName)
        {
            // Trim to length.
            proposedName = TrimToLength(proposedName);

            // Don't try to give it a name that already exists.
            bool success = false;
            string incrementedName = proposedName;

            while (!success)
            {
                try
                {
                    _worksheet.Name = incrementedName;
                    success = true;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    nameIncrement++;
                    incrementedName = proposedName + nameIncrement.ToString();
                }
            }

            return _worksheet;
        }

        private string TrimToLength(string proposedName)
        {
            // There's a 31-character limit.
            string cleanName = proposedName;

            if (proposedName.Length > MAX_LENGTH)
            {
                cleanName = proposedName.Substring(0, MAX_LENGTH);
            }

            return cleanName;
        }
    }
}
