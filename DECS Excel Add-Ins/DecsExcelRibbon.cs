using DECS_Excel_Add_Ins.Properties;
using Microsoft.Office.Core;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new DecsExcelRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Main class for DECS Excel Tools.
     *
     * The @c _GetImage methods specify the image used for each ribbon button.
     *
     * The @c On_ methods assign actions for each ribbon button push.
     *
     */
    [ComVisible(true)]
    public class DecsExcelRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DecsExcelRibbon() { }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c addRowsAcrossSheets button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap addRowsAcrossSheets_GetImage(IRibbonControl control)
        {
            return Resources.combine_sheets;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ImportList button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap buildListButton_GetImage(IRibbonControl control)
        {
            return Resources.clipboard;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ChopIntoTabsButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap chopIntoTabsButton_GetImage(IRibbonControl control)
        {
            return Resources.slice_into_tabs;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c PlotComparisonButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap comparisonPlotButton_GetImage(IRibbonControl control)
        {
            return Resources.comparison_graph;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c CopyFormatting button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap copyFormatButton_GetImage(IRibbonControl control)
        {
            return Resources.copy_formatting;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c CountWords button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap countWordsButton_GetImage(IRibbonControl control)
        {
            return Resources.abacus;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ConvertDates button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap dateConvertButton_GetImage(IRibbonControl control)
        {
            return Resources.calendar_with_gear;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c DateToText button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap dateToTextButton_GetImage(IRibbonControl control)
        {
            return Resources.calendar;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Deidentify button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap deidentifyButton_GetImage(IRibbonControl control)
        {
            return Resources.groucho;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Extend Timecard button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap extendTimecard_GetImage(IRibbonControl control)
        {
            return Resources.timecard;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ExtractMessage button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap extractMessageButton_GetImage(IRibbonControl control)
        {
            return Resources.nesting_dolls;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ExtractText button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap extractTextButton_GetImage(IRibbonControl control)
        {
            return Resources.uncorker;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c FormatResults button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap formatButton_GetImage(IRibbonControl control)
        {
            return Resources.paint_roller;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c HideDateTime button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap hideDateTimeButton_GetImage(IRibbonControl control)
        {
            return Resources.rubber_clock_small;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c HidePhysicianNames button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap hidePhysicianNamesButton_GetImage(IRibbonControl control)
        {
            return Resources.hide_identity;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c importSurvey button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap importSurveyButton_GetImage(IRibbonControl control)
        {
            return Resources.survey;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c lookupNpi button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap lookupNpiButton_GetImage(IRibbonControl control)
        {
            return Resources.NPI_Matching;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MatchPhysiciansButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap matchPhysiciansButton_GetImage(IRibbonControl control)
        {
            return Resources.match_people;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MatchTextButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap matchTextButton_GetImage(IRibbonControl control)
        {
            return Resources.match_pieces;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MergeRows button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap mergeNotesButton_GetImage(IRibbonControl control)
        {
            return Resources.merge_rows;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MergeRows button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap mergeRowsButton_GetImage(IRibbonControl control)
        {
            return Resources.combine_rows;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SetupConfig button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap notesConfigButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_setup_icon;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SearchNotes button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap notesSearchButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_search_icon;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c onCallList button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap onCallListButton_GetImage(IRibbonControl control)
        {
            return Resources.on_call;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SearchByEmail button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap searchByEmailButton_GetImage(IRibbonControl control)
        {
            return Resources.search_by_email;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SignalImportButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap signalImportButton_GetImage(IRibbonControl control)
        {
            return Resources.json;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SortTimes button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sortTimesButton_GetImage(IRibbonControl control)
        {
            return Resources.priority;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SortTimesSettings button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sortTimesSettingsButton_GetImage(IRibbonControl control)
        {
            return Resources.priority_settings;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Stripe button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap stripeButton_GetImage(IRibbonControl control)
        {
            return Resources.spreadsheet;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c AddSvi button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sviButton_GetImage(IRibbonControl control)
        {
            return Resources.ENV_EPHT_social;
        }

        /// <summary>
        /// When @c AddRowsAcrossSheets button is pressed, this method instantiates a @c RowAdder object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnAddRowsAcrossSheets(Office.IRibbonControl control)
        {
            RowAdder rowAdder = new RowAdder();
            rowAdder.Scan();
        }

        /// <summary>
        /// When @c AddSVI button is pressed, this method instantiates a @c SviProcessor object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnAddSVI(Office.IRibbonControl control)
        {
            SviProcessor sviProcessor = new SviProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            sviProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c BuildList button is pressed, instantiates a @c ListImporter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnBuildList(Office.IRibbonControl control)
        {
            ListImporter importer = new ListImporter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            importer.Scan(wksheet);
        }

        /// <summary>
        /// When @c ChopList button is pressed, instantiates a @c ListChopper object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnChopList(Office.IRibbonControl control)
        {
            ListChopper chopper = new ListChopper();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            chopper.Scan(wksheet);
        }

        /// <summary>
        /// When @c ComparisonPlot button is pressed, instantiates a @c ComparisonPlot object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnComparisonPlot(Office.IRibbonControl control)
        {
            ComparisonPlot plotter = new ComparisonPlot();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            plotter.Plot(wksheet);
        }

        /// <summary>
        /// When @c ConvertDates button is pressed, this method instantiates a @c MumpsDateConverter object & calls its @c ConvertColumn method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnConvertDates(Office.IRibbonControl control)
        {
            MumpsDateConverter converter = new MumpsDateConverter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.ConvertColumn(wksheet);
        }

        /// <summary>
        /// When @c CopyFormatting button is pressed, this method instantiates a @c Formatter object & calls its @c CopyFormat method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnCopyFormat(Office.IRibbonControl control)
        {
            Formatter formatter = new Formatter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            formatter.CopyFormat(wksheet);
        }

        /// <summary>
        /// When @c CountWords button is pressed, this method instantiates a @c WordCounter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnCountWords(Office.IRibbonControl control)
        {
            WordCounter wordCounter = new WordCounter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            wordCounter.Scan(wksheet);
        }

        /// <summary>
        /// When @c CopyFormatting button is pressed, this method instantiates a @c DateConverter object & calls its @c ToText method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnDatesToText(Office.IRibbonControl control)
        {
            DateConverter converter = new DateConverter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.ToText(wksheet);
        }

        /// <summary>
        /// When @c Deidentify button is pressed, this method instantiates a @c Deidentifier object & calls its @c GenerateHash method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnDeidentify(Office.IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.GenerateHash(wksheet);
        }

        /// <summary>
        /// When @c ExtendTimecard button is pressed, this method instantiates a @c Timecard object & calls its @c Extend method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnExtendTimecard(Office.IRibbonControl control)
        {
            Timecard timecard = new Timecard();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            timecard.Extend(wksheet);
        }

        /// <summary>
        /// When @c ExtractMessage button is pressed, instantiates a @c MessageUnpeeler object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnExtractMessage(Office.IRibbonControl control)
        {
            MessageUnpeeler unpeeler = new MessageUnpeeler();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            unpeeler.Scan(wksheet);
        }

        /// <summary>
        /// When @c ExtractMessage button is pressed, instantiates a @c MessageUnpeeler object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnExtractText(Office.IRibbonControl control)
        {
            TextExtractor extractor = new TextExtractor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            extractor.Extract(wksheet);
        }

        /// <summary>
        /// When @c FormatResults button is pressed, this method instantiates a @c Formatter object & calls its @c Format method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnFormat(Office.IRibbonControl control)
        {
            Formatter formatter = new Formatter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            formatter.Format(wksheet);
        }

        /// <summary>
        /// When @c hidePhysicianNames button is pressed, instantiates a @c Deidentifier object & calls its @c HidePhysicianNames method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnHidePhysicianNames(Office.IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.HidePhysicianNames(wksheet);
        }

        /// <summary>
        /// When @c lookupNpi button is pressed, instantiates a @c NpiLookup object & calls its @c Search method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnLookupNpi(Office.IRibbonControl control)
        {
            NpiLookup npiLookup = new NpiLookup();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            npiLookup.Search(wksheet);
        }

        /// <summary>
        /// When @c matchPhysicians button is pressed, instantiates a @c MatchPhysicians object & calls its @c Match method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMatchPhysicians(Office.IRibbonControl control)
        {
            PhysicianMatcher physicianMatcher = new PhysicianMatcher();
            physicianMatcher.Match();
        }

        /// <summary>
        /// When @c matchText button is pressed, instantiates a @c MatchText object & calls its @c Match method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMatchText(Office.IRibbonControl control)
        {
            TextMatcher textMatcher = new TextMatcher();
            textMatcher.Match();
        }

        /// <summary>
        /// When @c hideDateTime button is pressed, instantiates a @c Deidentifier object & calls its @c ObscureDateTime method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnObscureDateTime(Office.IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.ObscureDateTime(wksheet);
        }

        /// <summary>
        /// When @c importSurvey button is pressed, instantiates a @c SurveyResults object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnImportSurvey(Office.IRibbonControl control)
        {
            SurveyResults surveyProcessor = new SurveyResults();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            surveyProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c MergeNotes button is pressed, this method instantiates a @c MergeNotesForm.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMergeNotes(Office.IRibbonControl control)
        {
            MergeNotesForm form = new MergeNotesForm();
            form.Visible = true;
        }

        /// <summary>
        /// When @c MergeRows button is pressed, this method instantiates a @c MergeRowsForm.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMergeRows(Office.IRibbonControl control)
        {
            MergeRowsForm form = new MergeRowsForm();
            form.Visible = true;
        }

        /// <summary>
        /// When @c onCallList button is pressed, instantiates a @c OnCallListProcessor object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnProcessCallList(Office.IRibbonControl control)
        {
            OnCallListProcessor onCallListProcessor = new OnCallListProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            onCallListProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c SearchByEmail button is pressed, this method instantiates a @c EmailSearcher object & calls its @c Search method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchByEMail(Office.IRibbonControl control)
        {
            EmailSearcher emailSearcher = new EmailSearcher();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            emailSearcher.Search(wksheet);
        }

        /// <summary>
        /// When @c SetupConfig button is pressed, this method instantiates a @c DefineRulesForm object
        /// for the user to review & edit notes parsing rules.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchConfig(Office.IRibbonControl control)
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            NotesParser parser = new NotesParser(
                _worksheet: wksheet,
                withConfigFile: false,
                allRows: false
            );
            DefineRulesForm form = new DefineRulesForm(parser);
            form.Visible = true;
        }

        /// <summary>
        /// When @c SearchNotes button is pressed, this method instantiates a @c NotesParser object & calls its @c Parse method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchNotes(Office.IRibbonControl control)
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            NotesParser parser = new NotesParser(_worksheet: wksheet);
            parser.Parse();
        }

        /// <summary>
        /// When @c SignalImport button is pressed, instantiates a @c SignalTimeInNotes object & calls its @c Import method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSignalImport(Office.IRibbonControl control)
        {
            ImportSignalData parser = new ImportSignalData();
            parser.Import();
        }

        /// <summary>
        /// When @c SortTimes button is pressed, this method instantiates a @c TimeSorter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSortTimes(Office.IRibbonControl control)
        {
            TimeSorter timeSorter = new TimeSorter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            timeSorter.Scan(wksheet);
        }

        /// <summary>
        /// When @c SortTimesSettings button is pressed, this method instantiates a @c TimeSorterSettings object & calls its @c SetThresholds method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSortTimesSettings(Office.IRibbonControl control)
        {
            TimeSorterSettings setup = new TimeSorterSettings();
            setup.Set();
        }

        /// <summary>
        /// When @c stripe button is pressed, this method instantiates a @c Striper object & calls its @c Run method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnStripe(Office.IRibbonControl control)
        {
            Striper striper = new Striper();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            striper.Run(wksheet);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string txt = GetResourceText("DECS_Excel_Add_Ins.DecsExcelRibbon.xml");
            return txt;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (
                    string.Compare(
                        resourceName,
                        resourceNames[i],
                        StringComparison.OrdinalIgnoreCase
                    ) == 0
                )
                {
                    using (
                        StreamReader resourceReader = new StreamReader(
                            asm.GetManifestResourceStream(resourceNames[i])
                        )
                    )
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
