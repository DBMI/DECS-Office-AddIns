using DECS_Excel_Add_Ins.Properties;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
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
    [ComVisible(true)]
    public class DecsExcelRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DecsExcelRibbon() { }

        public Bitmap buildMrnsButton_GetImage(IRibbonControl control)
        {
            return Resources.clipboard;
        }

        public Bitmap dateConvertButton_GetImage(IRibbonControl control)
        {
            return Resources.calendar_with_gear;
        }

        public Bitmap formatButton_GetImage(IRibbonControl control)
        {
            return Resources.paint_roller;
        }

        public Bitmap notesConfigButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_setup_icon;
        }

        public Bitmap sviButton_GetImage(IRibbonControl control)
        {
            return Resources.ENV_EPHT_social;
        }

        public void OnAddSVI(Office.IRibbonControl control)
        {
            SviProcessor sviProcessor = new SviProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            sviProcessor.Scan(wksheet);
        }

        public void OnBuildMRN(Office.IRibbonControl control)
        {
            ListImporter importer = new ListImporter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            importer.Scan(wksheet);
        }

        public void OnConvertDates(Office.IRibbonControl control)
        {
            MumpsDateConverter converter = new MumpsDateConverter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.ConvertColumn(wksheet);
        }

        public void OnFormat(Office.IRibbonControl control)
        {
            Formatter formatter = new Formatter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            formatter.Format(wksheet);
        }

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

        public void OnSearchNotes(Office.IRibbonControl control)
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            NotesParser parser = new NotesParser(_worksheet: wksheet);
            parser.Parse();
        }

        public Bitmap notesSearchButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_search_icon;
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
