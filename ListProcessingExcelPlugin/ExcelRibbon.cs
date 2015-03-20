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
        private void CompareLists_Click(object sender, RibbonControlEventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="baseSheet">The sheet in which items will be bolded if they are not in the other sheet</param>
        /// <param name="minCol">The starting column in the range to compare (0 based)</param>
        /// <param name="maxCol">The ending column in the range to compare (0 based)</param>
        private void CompareLists(Excel.Worksheet baseSheet, long minCol, long maxCol)
        {

        }


    }
}
