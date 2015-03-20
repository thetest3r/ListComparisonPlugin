using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ListProcessingExcelPlugin
{
    public partial class ExcelRibbon
    {
        public Excel.Application ExcelApp
        {
            get
            {
                return (Marshal.GetActiveObject("Excel.Application") as Excel.Application);
            }
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // Get column arguments from the ribbon and send it off to compare
        private void CompareSheet1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet baseSheet = ExcelApp.Worksheets[1] as Excel.Worksheet, compareSheet = ExcelApp.Worksheets[2] as Excel.Worksheet;
            string minCol = "A", maxCol = "C";

            CompareLists(baseSheet, compareSheet, minCol, maxCol);
        }

        private void CompareSheet2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet baseSheet = ExcelApp.Worksheets[1], compareSheet = ExcelApp.Worksheets[0];
            string minCol = "A", maxCol = "C";

            CompareLists(baseSheet, compareSheet, minCol, maxCol);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="baseSheet">The sheet in which items will be bolded if they are not in the other sheet</param>
        /// <param name="minCol">The starting column in the range to compare (0 based)</param>
        /// <param name="maxCol">The ending column in the range to compare (0 based)</param>
        private void CompareLists(Excel.Worksheet baseSheet, Excel.Worksheet compareSheet, string minCol, string maxCol)
        {
            MessageBox.Show("Comparing " + baseSheet.Name + " to " + compareSheet.Name);
        }

        


    }
}
