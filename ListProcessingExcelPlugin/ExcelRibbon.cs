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
using System.Text.RegularExpressions;

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

        private void minColEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            SelectColumnsInRange();

            
            //Excel.Range firstRow = activeWorksheet.get_Range("A1");
            //firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            //Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            //newFirstRow.Value2 = "This text was added by using code";
        }

        private void maxColEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            SelectColumnsInRange();
        }

        private void SelectColumnsInRange()
        {
            Excel.Worksheet activeWorksheet = ExcelApp.ActiveSheet as Excel.Worksheet;

            string minCol = minColEditBox.Text;
            string maxCol = maxColEditBox.Text;

            if (ValidateColumnInput(minCol, maxCol, false))
            {
                Range rng = activeWorksheet.get_Range(minCol + "1", maxCol + "1");
                rng.EntireColumn.Select();
            }
        }

        // Get column arguments from the ribbon and send it off to compare
        private void CompareSheet1_Click(object sender, RibbonControlEventArgs e)
        {
            string minCol = minColEditBox.Text, maxCol = maxColEditBox.Text;

            if (ValidateSheets() && ValidateColumnInput(minCol, maxCol, true))
            {
                Excel.Worksheet baseSheet = ExcelApp.Worksheets[1] as Excel.Worksheet;
                Excel.Worksheet compareSheet = ExcelApp.Worksheets[2] as Excel.Worksheet;
                CompareLists(baseSheet, compareSheet, minCol, maxCol);
            }
        }

        private void CompareSheet2_Click(object sender, RibbonControlEventArgs e)
        {
            string minCol = minColEditBox.Text, maxCol = maxColEditBox.Text;

            if (ValidateSheets() && ValidateColumnInput(minCol, maxCol, true))
            {
                Excel.Worksheet baseSheet = ExcelApp.Worksheets[1] as Excel.Worksheet;
                Excel.Worksheet compareSheet = ExcelApp.Worksheets[2] as Excel.Worksheet;
                CompareLists(baseSheet, compareSheet, minCol, maxCol);
            }
        }

        private bool ValidateSheets()
        {
            // Make sure the user has two sheets to compare with
            if (ExcelApp.Worksheets.Count < 2)
            {
                MessageBox.Show("You must have at least two sheets. The first two sheets should contain your lists.");
                return false;
            }

            return true;
        }

        public bool ValidateColumnInput(string minCol, string maxCol, bool displayErrorMessages)
        {
            // Verify that the column inputs contain only characters
            bool minColIllegal = Regex.IsMatch(minCol, "[^a-z|A-Z]");
            bool maxColIllegal = Regex.IsMatch(maxCol, "[^a-z|A-Z]");

            if (minColIllegal || maxColIllegal)
            {
                if (displayErrorMessages)
                    MessageBox.Show("The column inputs can only contain letters that represent columns");
                return false;
            }
            

            // Verify that the minimum column is less than the maximum column
            int comparisonResult = minCol.CompareTo(maxCol); // Compare yields -1 is less than, 0 if equal, 1 if greater than
            
            if (comparisonResult != -1)
            {
                if (displayErrorMessages)
                    MessageBox.Show("The minimum column range must be less than the maximum column range");
                return false;
            }
                
            // Input is correct
            return true;
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

            //Gets the range and columns in the worksheet that are used. Range will be used to loop, and col to keep data intact
            //TODO: Potential documentation that will assist in helping out
            //https://support.microsoft.com/en-us/kb/302096
            var baseTotalNumOfCols = baseSheet.UsedRange.Columns.Count;
            var baseTotalNumOfRows = baseSheet.UsedRange.Columns.Count;

            var compareTotalNumOfCols = compareSheet.UsedRange.Columns.Count;
            var compareTotalNumOfRows = compareSheet.UsedRange.Columns.Count;



        }

        

        


    }
}
