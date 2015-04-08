using System;
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
            //this.Application.WorkbookOpen += Application_WorkbookOpen;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        public void GetApplicationInstance()
        {
            
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            Excel.Worksheet sheet1 = this.Application.Worksheets[1];
            Globals.Ribbons.Ribbon1.sheet1RangeBox.Label = sheet1.Name + " Range";

            if (this.Application.Worksheets.Count > 1)
            {
                Excel.Worksheet sheet2 = this.Application.Worksheets[2];
                Globals.Ribbons.Ribbon1.sheet2RangeBox.Label = sheet2.Name + " Range";
            }
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
