using DocumentFormat.OpenXml.Bibliography;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class to create a histogram from Excel data.
     */
    internal class HistogramBuilder
    {
        // Like the census tract number.
        private Range categoryColumn = null;

        // Optional. Maybe this is a column that says whether patients are overdue or not.
        private Range scoreColumn = null;

        internal void Build(Worksheet worksheet)
        {
            if (SelectColumns(worksheet))
            {
                Dictionary<string, int> counts = CountCells();
                BuildHistogram(counts);
            }
        }

        private void BuildHistogram(Dictionary<string, int> counts)
        {
            // Create & set up histogram sheet.
            Worksheet histogramSheet = Utilities.CreateNewNamedSheet("Histogram");
            Range target = (Range)histogramSheet.Cells[1, 1];
            target.Value = categoryColumn.Value.ToString();
            target.Offset[0, 1].Value = "Number";

            if (scoreColumn != null)
            {
                target.Offset[0, 1].Value += " " + scoreColumn.Value.ToString();
            }

            // Sort by value descending, then by key ascending for tie-breaking.
            var sortedItems = counts.OrderByDescending(pair => pair.Value)
                                  .ThenBy(pair => pair.Key);
            int rowOffset = 1;

            foreach (var item in sortedItems)
            {
                target.Offset[rowOffset, 0].Value = item.Key;
                string cellAddress = target.Offset[rowOffset, 0].AddressLocal.ToString();
                target.Offset[rowOffset, 1].Value = item.Value;
                rowOffset++;
            }
        }

        private int ConvertScoreToIncrement(string scoreValue)
        {
            int scoreIncrement = 0;

            if (!string.IsNullOrEmpty(scoreValue))
            {
                switch (scoreValue)
                {
                    case "0":
                    case "N":
                        break;

                    case "1":
                    case "Y":
                        scoreIncrement = 1;
                        break;

                    default:
                        break;
                }
            }

            return scoreIncrement;
        }

        private Dictionary<string, int> CountCells()
        {
            Dictionary<string, int> counts = new Dictionary<string, int>();
            int rowOffset = 1;
            int numConsecutiveFailures = 0;
            string categoryValue = string.Empty;
            int scoreIncrement = 1;

            while (true) 
            {
                try
                {
                    categoryValue = categoryColumn.Offset[rowOffset, 0].Value.ToString();

                    // We're looking for "0" or "1" in the score column.
                    // But maybe it's "Y" or "N"?
                    if (scoreColumn != null)
                    {
                        scoreIncrement = ConvertScoreToIncrement(scoreColumn.Offset[rowOffset, 0].Value.ToString());
                    }

                    if (counts.ContainsKey(categoryValue))
                    {
                        counts[categoryValue] += scoreIncrement;
                    }
                    else
                    {
                        counts[categoryValue] = scoreIncrement;
                    }
                }
                catch (RuntimeBinderException)
                {
                    numConsecutiveFailures++;
                }

                rowOffset++;

                // An occasional miss is ok, but three in a row & we've run outta data.
                if (numConsecutiveFailures >= 3)
                {
                    break;
                }
            }

            return counts;
        }

        private bool SelectColumns(Worksheet worksheet) 
        {
            bool success = false;
            List<string> columnNames = Utilities.GetColumnNames(worksheet);

            using (ChooseHistogramColumnsForm form = new ChooseHistogramColumnsForm(columnNames))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    categoryColumn = Utilities.TopOfNamedColumn(worksheet, form.categoryColumn);

                    if (!string.IsNullOrEmpty(form.scoreColumn))
                    {
                        scoreColumn = Utilities.TopOfNamedColumn(worksheet, form.scoreColumn);
                    }

                    success = true;
                }
                else if (result == DialogResult.Cancel)
                {
                    // Then we're done here.
                    return success;
                }
            }

            return success;
        }
    }
}
