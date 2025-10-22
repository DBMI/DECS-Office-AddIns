using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class RowAdder
    {
        private Application application;
        private Dictionary<string, int> comboDepartments;
        private List<Worksheet> sheets;
        private List<int> validColOffsets;  // The column offsets for "Num" columns (not "%").

        internal RowAdder()
        {
            application = Globals.ThisAddIn.Application;
        }

        internal void Scan()
        {
            // All the individual sheets.
            sheets = Utilities.CollectAllWorksheets((Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook);

            // The sum sheet.
            Worksheet comboSheet = PrepareComboSheet(sheets[0]);

            foreach (Worksheet sheet in sheets)
            {
                AddRowsFromSheet(sheet, comboSheet);
            }
        }

        private void AddRowsFromSheet(Worksheet sourceSheet, Worksheet targetSheet)
        {
            Range sourceRng = (Range)sourceSheet.Cells[1, 1];
            Range targetRng = (Range)targetSheet.Cells[1, 1];

            // Skip the header row.
            int sourceRowOffset = 1;

            while (true)
            {
                try
                {
                    string departmentName = sourceRng.Offset[sourceRowOffset, 0].Value;

                    // Have we reached the end of the page?
                    if (departmentName is null)
                    {
                        break;
                    }

                    if (!comboDepartments.ContainsKey(departmentName))
                    {
                        InsertNewDepartment(departmentName, targetSheet);
                    }

                    int targetRowOffset = comboDepartments[departmentName];

                    AddThisDeptNumbers(sourceRng.Offset[sourceRowOffset], targetRng.Offset[targetRowOffset]);
                    sourceRowOffset++;
                }
                catch (NullReferenceException)
                {
                    break;
                }
            }
        }

        private void AddThisDeptNumbers(Range sourceRange, Range targetRange)
        {
            foreach (int colOffset in validColOffsets)
            { 
                double sourceValue = sourceRange.Offset[0, colOffset].Value;
                double targetValue = 0;

                try
                {
                    targetValue = targetRange.Offset[0, colOffset].Value;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                }

                targetRange.Offset[0, colOffset].Value = targetValue + sourceValue;
            }
        }

        private void CopyHeader(Worksheet sourceSheet, Worksheet targetSheet)
        {
            Range sourceRng = (Range)sourceSheet.Cells[1, 1];
            Range targetRng = (Range)targetSheet.Cells[1, 1];

            int colOffset = 0;
            validColOffsets = new List<int>();

            while (true)
            {
                try
                {
                    string sourceData = sourceRng.Offset[0, colOffset].Value;

                    // Have we copied them all?
                    if (sourceData is null)
                    {
                        break;
                    }

                    // Keep track of which columns are to be added ("Num messages")
                    // and which can't be added ("% of messages").
                    if (sourceData.Contains("Num") && !sourceData.Contains("%"))
                    {
                        validColOffsets.Add(colOffset);
                    }

                    targetRng.Offset[0, colOffset].Value = sourceData;
                    colOffset++;
                }
                catch (NullReferenceException)
                {
                    break;
                }
            }
        }

        private void InsertNewDepartment(string newName, Worksheet targetSheet)
        {
            int maxOffsetSoFar = 0;

            try
            {
                maxOffsetSoFar = comboDepartments.Values.Max();
            }
            catch (InvalidOperationException) { }
            
            int newOffset = maxOffsetSoFar + 1;
            comboDepartments.Add(newName, newOffset);

            // Copy over the department name.
            Range targetRng = (Range)targetSheet.Cells[1, 1];
            targetRng.Offset[newOffset, 0].Value = newName;
        }

        private Worksheet PrepareComboSheet(Worksheet sourceSheet)
        {
            // Create combo sheet.
            Worksheet comboSheet = Utilities.CreateNewNamedSheet("Combined");

            // Copy header to combo sheet.
            CopyHeader(sourceSheet, comboSheet);

            // Initialize dictionary of departments to row numbers.
            comboDepartments = new Dictionary<string, int>();

            return comboSheet;
        }
    }
}
