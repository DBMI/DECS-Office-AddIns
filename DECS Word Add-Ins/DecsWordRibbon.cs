using DecsWordAddIns.Properties;
using Microsoft.Office.Core;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new DecsWordRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

namespace DecsWordAddIns
{
    /**
     * @brief Main class for DECS Excel Tools.
     *
     * The @c _GetImage methods specify the image used for each ribbon button.
     *
     * The @c On_ methods assign actions for each ribbon button push.
     */
    [ComVisible(true)]
    public class DecsWordRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DecsWordRibbon() { }

        /// <summary>
        /// Defines what happens when @c ImportList button is pushed.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        public void OnImportList(Office.IRibbonControl control)
        {
            ListImporter importer = new ListImporter();
            importer.Scan(Globals.ThisAddIn.Application.ActiveDocument);
        }

        /// <summary>
        /// Defines what happens when @c ExtractICD button is pushed.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        public void OnExtractICD(Office.IRibbonControl control)
        {
            IcdExtractor extractor = new IcdExtractor();
            extractor.Scan(Globals.ThisAddIn.Application.ActiveDocument);
        }

        /// <summary>
        /// Defines what happens when @c formatMsg button is pushed.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        public void OnFormatMsg(Office.IRibbonControl control)
        {
            MsgFormatter formatter = new MsgFormatter();
            formatter.Format(Globals.ThisAddIn.Application.ActiveDocument);
        }

        /// <summary>
        /// Defines what happens when @c SetupProject button is pushed.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        public void OnParseSOW(Office.IRibbonControl control)
        {
            ScopeOfWorkParser parser = new ScopeOfWorkParser();
            parser.SetupProject(Globals.ThisAddIn.Application.ActiveDocument);
        }

        /// <summary>
        /// Supplies the format_messages image to the @c formatMsg button.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        /// <returns></returns>
        public Bitmap formatMsgButton_GetImage(IRibbonControl control)
        {
            return Resources.format_messages;
        }

        /// <summary>
        /// Supplies the icd_10_zoom image to the @c ExtractICD button.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        /// <returns></returns>
        public Bitmap icdExtractButton_GetImage(IRibbonControl control)
        {
            return Resources.icd_10_zoom;
        }

        /// <summary>
        /// Supplies the clipboard image to the @c ImportList button.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        /// <returns></returns>
        public Bitmap importListButton_GetImage(IRibbonControl control)
        {
            return Resources.clipboard;
        }

        /// <summary>
        /// Supplies the crane image to the @c SetupProject button.
        /// </summary>
        /// <param name="control">The @c IRibbon object</param>
        /// <returns></returns>
        public Bitmap scopeOfWorkParserButton_GetImage(IRibbonControl control)
        {
            return Resources.crane;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("DecsWordAddIns.DecsWordRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
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
