using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class Rows
    {
        private int? startRow = null;
        private int? endRow = null;

        internal Rows() { }

        internal Rows(int _start, int _end)
        {
            startRow = _start;
            endRow = _end;
        }
        internal int End() { return endRow.Value; }
        internal bool HasEnd() { return endRow.HasValue; }
        internal bool HasStart() { return startRow.HasValue; }
        internal void SetEnd(int _end)
        {
            endRow = _end;
        }
        internal void SetStart(int _start)
        {
            startRow = _start;
        }
        internal Range SpecifyColumn(Worksheet sheet, int columnNum)
        {
            return sheet.Range[sheet.Cells[startRow, columnNum], sheet.Cells[endRow, columnNum]];
        }
        internal int Start() { return startRow.Value; }

        internal bool Valid()
        {
            return startRow.HasValue && endRow.HasValue;
        }
    }

    internal class ComparisonPlot
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private string name1ColumnName;
        private string name2ColumnName;
        private string sheet1Name;
        private string sheet2Name;
        private string time1ColumnName;
        private string time2ColumnName;
        private string value1ColumnName;
        private string value2ColumnName;
        private Dictionary<string, Worksheet> worksheets;

        internal ComparisonPlot()
        {
            application = Globals.ThisAddIn.Application;
        }

        internal void Plot(Worksheet thisSheet)
        {
            Worksheet tempSheet = SetDefaultChartType(thisSheet);

            Workbook workbook = thisSheet.Parent;
            worksheets = Utilities.GetWorksheets();

            // Ask user to specify which columns to plot.
            if (GetUserPreferences(worksheets))
            {
                // Get list of unique names from first plot source.
                List<string> names1 = GetNames(sheet1Name, name1ColumnName);
                names1.Reverse(); // So that the sheets will be in alpha order left -> right
                Worksheet sheet1 = worksheets[sheet1Name];
                Dictionary<string, Range> sheet1Columns = Utilities.GetColumnRangeDictionary(sheet1);
                Range time1Column = sheet1Columns[time1ColumnName];
                int time1ColumnNum = time1Column.Column;
                Range value1Column = sheet1Columns[value1ColumnName];
                int value1ColumnNum = value1Column.Column;

                // Get list of unique names from second plot source.
                List<string> names2 = GetNames(sheet2Name, name2ColumnName);
                Worksheet sheet2 = worksheets[sheet2Name];
                Dictionary<string, Range> sheet2Columns = Utilities.GetColumnRangeDictionary(sheet2);
                Range time2Column = sheet2Columns[time2ColumnName];
                int time2ColumnNum = time2Column.Column;
                Range value2Column = sheet2Columns[value2ColumnName];
                int value2ColumnNum = value2Column.Column;

                foreach (string thisName in names1)
                {
                    // Which rows in names columns match this name?
                    Rows rows1 = Utilities.FindNamesExact(sheet1, name1ColumnName, thisName);

                    if (!rows1.Valid()) { continue; }

                    // Turn row start, end into Ranges for plotting.
                    Range times1 = rows1.SpecifyColumn(sheet1, time1ColumnNum);
                    Range values1 = rows1.SpecifyColumn(sheet1, value1ColumnNum);
                    Range range1 = MergeRangesWithoutNulls(times1, values1);

                    if (range1 is null)
                    {
                        continue;
                    }

                    // Which name in sheet 2 most closely matches this name from sheet 1?
                    NameComparison nameComparison = new NameComparison(thisName);
                    string thisNameInSheet2 = nameComparison.FindBestMatch(names2, maxDistanceAllowed: 0.1);

                    if (!string.IsNullOrEmpty(thisNameInSheet2))
                    {
                        // Which rows in sheet 2's names columns match this name?
                        Rows rows2 = Utilities.FindNamesExact(sheet2, name2ColumnName, thisNameInSheet2);

                        if (rows2.Valid())
                        {
                            // Initialize chart.
                            Chart chartSheet = workbook.Charts.Add();
                            chartSheet.Location(XlChartLocation.xlLocationAsNewSheet, MergeNameWithExtra("plot", thisName));
                            chartSheet.SetSourceData(range1);
                            chartSheet.HasTitle = true;
                            chartSheet.ChartTitle.Text = thisName;
                            Series series1 = chartSheet.SeriesCollection(1);
                            series1.Name = MergeNameWithExtra(sheet1Name, thisName);
                            chartSheet.ChartType = XlChartType.xlXYScatterLines;
                            Axis xAxis = chartSheet.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                            xAxis.MajorGridlines.Border.Color = System.Drawing.Color.Gray;
                            Axis yAxis = chartSheet.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                            yAxis.MajorGridlines.Border.Color = System.Drawing.Color.Gray;
                            chartSheet.Select();

                            Range times2 = rows2.SpecifyColumn(sheet2, time2ColumnNum);
                            Range values2 = rows2.SpecifyColumn(sheet2, value2ColumnNum);

                            Series newSeries = chartSheet.SeriesCollection().NewSeries();
                            newSeries.Name = MergeNameWithExtra(sheet2Name, thisNameInSheet2);
                            newSeries.Values = values2;
                            newSeries.XValues = times2;
                        }
                    }
                }
            }

            application.DisplayAlerts = false;
            tempSheet.Delete();
            application.DisplayAlerts = true;
        }

        private List<string> GetNames(string sheetName, string columnName)
        {
            Worksheet sheet = worksheets[sheetName];
            List<string> names = Utilities.ExtractColumnUnique(sheet, columnName);
            return names;
        }

        private bool GetUserPreferences(Dictionary<string, Worksheet> worksheets)
        {
            bool success = false;

            using (PlotSelectionForm form = new PlotSelectionForm(worksheets))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    name1ColumnName = form.name1Column;
                    name2ColumnName = form.name2Column;
                    sheet1Name = form.sheet1Name;
                    sheet2Name = form.sheet2Name;
                    time1ColumnName = form.time1Column;
                    time2ColumnName = form.time2Column;
                    value1ColumnName = form.value1Column;
                    value2ColumnName = form.value2Column;
                    success = true;
                }
            }

            return success;
        }

        private string MergeNameWithExtra(string extraName, string name)
        {
            int maxNameLength = 31;
            string seriesName = name + " " + extraName;

            if (seriesName.Length > maxNameLength)
            {
                seriesName = seriesName.Substring(0, maxNameLength);
            }

            return seriesName;
        }

        private Range MergeRangesWithoutNulls(Range times, Range values)
        {
            Range combo = null;
            Range timesValid = null;
            Range valuesValid = null;

            for (int counter = 1; counter < times.Cells.Count; counter++)
            {
                Range thisTimesCell = times.Cells[counter];
                Range thisValuesCell = values.Cells[counter];
                string cell_contents;

                try
                {
                    cell_contents = Convert.ToString(thisTimesCell.Value2);

                    if (!string.IsNullOrEmpty(cell_contents) && cell_contents.ToUpper() != "NULL")
                    {
                        if (timesValid is null)
                        {
                            timesValid = thisTimesCell;
                            valuesValid = thisValuesCell;

                        }
                        else
                        {
                            timesValid = application.Union(timesValid, thisTimesCell);
                            valuesValid = application.Union(valuesValid, thisValuesCell);
                        }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            }

            if (timesValid != null && valuesValid != null)
            {
                combo = application.Union(timesValid, valuesValid);
            }

            return combo;
        }

        private Worksheet SetDefaultChartType(Worksheet worksheet)
        {
            Worksheet tempSheet = Utilities.CreateNewNamedSheet(worksheet, "temp");
            ChartObjects chartObjects = (ChartObjects)tempSheet.ChartObjects();
            ChartObject tempChartObject = chartObjects.Add(100, 100, 300, 200);
            Chart tempChart = tempChartObject.Chart;
            tempChart.SetDefaultChart(XlChartType.xlXYScatterLines);
            return tempSheet;
        }
    }
}
