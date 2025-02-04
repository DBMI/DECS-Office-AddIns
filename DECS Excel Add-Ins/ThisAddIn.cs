﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using log4net;

namespace DECS_Excel_Add_Ins
{
    public partial class ThisAddIn
    {
        //  https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-a-custom-tab-by-using-ribbon-xml?view=vs-2022&tabs=csharp
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DecsExcelRibbon();
        }

        // https://stackoverflow.com/a/28546547/18749636
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            log4net.Config.XmlConfigurator.Configure();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
