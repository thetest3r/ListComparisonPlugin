using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ListProcessingExcelPlugin
{
    public partial class ExcelRibbon
    {
        public Excel._Application ExcelApp
        {
            get
            {
                return (Marshal.GetActiveObject("Excel.Application") as Excel._Application);
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ProcessList_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the worksheet
            Excel._Worksheet activeWorksheet = ExcelApp.ActiveSheet;

            // Get the first row and move it down to make room for the new line of text
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Get the new first row and put text in it
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = "This text was added by using code";
        }
    }
}
