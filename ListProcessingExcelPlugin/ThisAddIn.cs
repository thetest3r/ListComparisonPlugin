﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


//Reference for Development
//http://support.microsoft.com/kb/302901
//https://msdn.microsoft.com/en-us/library/aa289518(v=vs.71).aspx


namespace ListProcessingExcelPlugin
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            ExcelRibbon.LoadButtonNames();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
