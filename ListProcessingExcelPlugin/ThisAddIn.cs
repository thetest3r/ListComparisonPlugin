using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;


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
            this.Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            this.Application.WorkbookOpen += Application_WorkbookOpen;
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            Wb.SheetActivate += Wb_SheetActivate;
        }

        void Wb_SheetActivate(object Sh)
        {
            string test = "testing";
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }


        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            ExcelRibbon.RepopulateSheetDropDowns();

            //var sheet1DropDown = Globals.Ribbons.Ribbon1.sheet1DropDown;
            //var sheet2DropDown = Globals.Ribbons.Ribbon1.sheet2DropDown;

            //sheet1DropDown.Items.Clear();
            //sheet2DropDown.Items.Clear();

            //foreach (Excel.Worksheet sheet in this.Application.Worksheets)
            //{
            //    RibbonDropDownItem item1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem(), item2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            //    item1.Label = sheet.Name;
            //    item2.Label = sheet.Name;

            //    sheet1DropDown.Items.Add(item1);
            //    sheet2DropDown.Items.Add(item2);
            //}

            //sheet1DropDown.SelectedItemIndex = 0;
            //sheet2DropDown.SelectedItemIndex = (sheet2DropDown.Items.Count > 1) ? 1 : -1;
        }



        void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            ExcelRibbon.RepopulateSheetDropDowns();

            //var sheet1DropDown = Globals.Ribbons.Ribbon1.sheet1DropDown;
            //var sheet2DropDown = Globals.Ribbons.Ribbon1.sheet2DropDown;

            //RibbonDropDownItem item1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem(), item2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            //item1.Label = (Sh as Excel.Worksheet).Name;
            //item2.Label = (Sh as Excel.Worksheet).Name;

            //sheet1DropDown.Items.Add(item1);
            //sheet2DropDown.Items.Add(item2);
        }

        void ThisWorkbook_SheetActivate(object Sh)
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
