using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace DECS_Excel_Add_Ins
{
    internal class FileMerger
    {
        private Dictionary<string, int> columnIndices;
        private bool firstFile = true;
        private string folder = String.Empty;
        private Range target;
        private const int widthDateTimeColumn = 20;
        private const int widthIdColumn = 10;
        private const int widthMetricNameColumn = 19;
        private const int widthMetricDescripColumn = 25;
        private const int widthEncounterColumn = 11;

        internal FileMerger()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                // Because we don't specify an opening directory,
                // the dialog will open in the last directory used.
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified folder.
                    folder = folderDialog.SelectedPath;
                }
            }

            columnIndices = new Dictionary<string, int>();
        }

        private void DiscoverColumnStartIndices(string line)
        {
            columnIndices["NOTE_ID"] = line.IndexOf("NOTE_ID");
            columnIndices["EFF_LOCAL_DTTM"] = line.IndexOf("EFF_LOCAL_DTTM");
            columnIndices["METRIC_NAME"] = line.IndexOf("METRIC_NAME");
            columnIndices["METRIC_DESC"] = line.IndexOf("METRIC_DESC");
            columnIndices["PAT_ENC_CSN_ID"] = line.IndexOf("PAT_ENC_CSN_ID");
        }

        private void LabelFile()
        {
            target.Offset[0, 0].Value = "NOTE_ID";
            target.Offset[0, 1].Value = "EFF_LOCAL_DTTM";
            target.Offset[0, 2].Value = "METRIC_NAME";
            target.Offset[0, 3].Value = "METRIC_DESC";
            target.Offset[0, 4].Value = "PAT_ENC_CSN_ID";
            target = target.Offset[1, 0];
        }

        internal void Merge(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            if (!string.IsNullOrEmpty(folder))
            {
                target = (Range)worksheet.Cells[1, 1];

                // Find all the .csv files in the folder.
                List<string> csvFiles = Directory.GetFiles(folder, "*.csv").ToList<string>();

                foreach (string csvFile in csvFiles) 
                {
                    ReadFile(csvFile);
                }
            }
        }

        internal void ReadFile(string path)
        {
            // Read lines lazily.
            IEnumerable<string> lines = File.ReadLines(path);
            bool foundPayload = false;

            // Process each line in a loop.
            foreach (string line in lines)
            {
                if (!foundPayload) 
                {
                    if (line.StartsWith("NOTE_ID"))
                    {
                        foundPayload = true;

                        if (firstFile)
                        {
                            DiscoverColumnStartIndices(line);
                            LabelFile();
                            firstFile = false;
                        }
                    }

                    continue;
                }

                if (line.StartsWith("-----"))
                {
                    continue;
                }

                try 
                {
                    target.Offset[0, 0].Value2 = line.Substring(columnIndices["NOTE_ID"], widthIdColumn);
                    target.Offset[0, 1].Value2 = line.Substring(columnIndices["EFF_LOCAL_DTTM"], widthDateTimeColumn);
                    target.Offset[0, 2].Value2 = line.Substring(columnIndices["METRIC_NAME"], widthMetricNameColumn);
                    target.Offset[0, 3].Value2 = line.Substring(columnIndices["METRIC_DESC"], widthMetricDescripColumn);
                    target.Offset[0, 4].Value2 = line.Substring(columnIndices["PAT_ENC_CSN_ID"], widthEncounterColumn);
                    target = target.Offset[1, 0];
                }
                catch (ArgumentOutOfRangeException)
                {
                    // Ran out of data.
                    return;
                }                
            }
        }
    }
}
