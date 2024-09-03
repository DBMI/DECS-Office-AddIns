using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class Deidentifier
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private int lastRow;
        private string sSourceData;
        private byte[] tmpSource;
        private byte[] tmpHash;

        internal Deidentifier()
        {
            application = Globals.ThisAddIn.Application;
        }

        // https://learn.microsoft.com/en-us/troubleshoot/developer/visualstudio/csharp/language-compilers/compute-hash-values
        private static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);

            for (i = 0; i < arrInput.Length; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }

            return sOutput.ToString();
        }

        internal void GenerateHash(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            // Any columns selected?
            List<Range> selectedColumns = Utilities.GetSelectedCols(application, lastRow);

            if (selectedColumns.Count == 0)
            {
                MessageBox.Show("Please select column(s) that represent a unique identifier.");
                return;
            }

            // Make room for new column.
            Range lastSourceColumn = selectedColumns.Last();
            Range hashColumn = Utilities.InsertNewColumn(range: lastSourceColumn, newColumnName: "Scrambled Identifier", side: InsertSide.Right);

            string sourceData;
            Range target;
            byte[] tmpHash;
            byte[] tmpSource;

            for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
            {
                target = (Range)worksheet.Cells[rowNumber, hashColumn.Column];
                sourceData = Utilities.CombineColumns(worksheet, rowNumber, selectedColumns);

                if (!string.IsNullOrEmpty(sourceData))
                {
                    // Create a byte array from source data.
                    tmpSource = ASCIIEncoding.ASCII.GetBytes(sourceData);

                    // Initialize a SHA256 hash object.
                    using (SHA256 mySHA256 = SHA256.Create())
                    {
                        tmpHash = mySHA256.ComputeHash(tmpSource);
                        target.Value = ByteArrayToString(tmpHash);
                    }
                }
            }
        }
    }
}