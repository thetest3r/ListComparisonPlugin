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
            // Get a list of the sheets


            // Change the buttons' labels to the first two sheets names


        }


        private void minColEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            SelectColumnsInRange();
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
            bool headerRow = headerRowCheckBox.Checked;

            if (ValidateSheets() && ValidateColumnInput(minCol, maxCol, true))
            {
                Excel.Worksheet baseSheet = ExcelApp.Worksheets[1] as Excel.Worksheet;
                Excel.Worksheet compareSheet = ExcelApp.Worksheets[2] as Excel.Worksheet;
                List<int> sheet1Indices = CompareLists(baseSheet, compareSheet, minCol, maxCol, headerRow);
                DeleteUnmatchedRows(baseSheet, sheet1Indices, headerRow);
            }
        }

        private void CompareSheet2_Click(object sender, RibbonControlEventArgs e)
        {
            string minCol = minColEditBox.Text, maxCol = maxColEditBox.Text;
            bool headerRow = headerRowCheckBox.Checked;

            if (ValidateSheets() && ValidateColumnInput(minCol, maxCol, true))
            {
                Excel.Worksheet baseSheet = ExcelApp.Worksheets[2] as Excel.Worksheet;
                Excel.Worksheet compareSheet = ExcelApp.Worksheets[1] as Excel.Worksheet;
                List<int> sheet2Indices = CompareLists(baseSheet, compareSheet, minCol, maxCol, headerRow);
                DeleteUnmatchedRows(baseSheet, sheet2Indices, headerRow);
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

        private bool ValidateColumnInput(string minCol, string maxCol, bool displayErrorMessages)
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

            if (comparisonResult == 1)
            {
                if (displayErrorMessages)
                    MessageBox.Show("The minimum column range must be less than or equal to the maximum column range");
                return false;
            }

            // Input is correct
            return true;
        }

        /// <summary>
        /// Compare every row in the base sheet to every row in the compare sheet and looks for matches. The users parameters determine exactly what is compared
        /// </summary>
        /// <param name="baseSheet">The sheet in which items will be bolded if they are not in the other sheet</param>
        /// <param name="minCol">The starting column in the range to compare (0 based)</param>
        /// <param name="maxCol">The ending column in the range to compare (0 based)</param>
        /// <returns>Returns a list of row indices in the base sheet that have matches in the compare sheet</returns>
        private List<int> CompareLists(Excel.Worksheet baseSheet, Excel.Worksheet compareSheet, string minCol, string maxCol, bool headerRow)
        {
            List<int> baseSheetIndices = new List<int>();

            //Gets the range and columns in the worksheet that are used. Range will be used to loop, and col to keep data intact
            var baseTotalNumOfCols = baseSheet.UsedRange.Columns.Count;
            var baseTotalNumOfRows = baseSheet.UsedRange.Rows.Count;

            var compareTotalNumOfCols = compareSheet.UsedRange.Columns.Count;
            var compareTotalNumOfRows = compareSheet.UsedRange.Rows.Count;

            // Compare each row in the base sheet to every row in the compare sheet
            // If a match is found, save the current row's index
            for (int i = (!headerRow) ? 1 : 2; i <= baseTotalNumOfRows; i++)
            {
                StringBuilder sb = new StringBuilder();

                // Create the comparison string for the row in the base sheet
                for (int j = 1; j <= baseTotalNumOfCols; j++)
                {
                    string colName = GetColNameFromIndex(j);

                    // Make sure to not go past the max column
                    if (colName.CompareTo(maxCol.ToUpper()) <= 0)
                    {
                        // Add each cells' contents to the string
                        Range cell = baseSheet.Cells[i, j] as Range;
                        sb.Append(Convert.ToString(cell.Value).Trim());
                    }
                    else
                        break;
                }

                // Compare the row in the base sheet with every row in the compare sheet
                for (int k = (!headerRow) ? 1 : 2; k <= compareTotalNumOfRows; k++)
                {
                    StringBuilder sb1 = new StringBuilder();

                    // Create the comparison string for the row in the compare sheet
                    for (int j = 1; j <= compareTotalNumOfCols; j++)
                    {
                        string colName = GetColNameFromIndex(j);

                        // Make sure to not go past the max column
                        if (colName.CompareTo(maxCol.ToUpper()) <= 0)
                        {
                            // Add each cells' contents to the string
                            Range cell = compareSheet.Cells[k, j] as Range;
                            sb1.Append(Convert.ToString(cell.Value).Trim());
                        }
                        else
                            break;
                    }

                    // Compare the two rows and see if they are the same
                    if (sb.ToString() == sb1.ToString())
                    {
                        baseSheetIndices.Add(i);
                        break;
                    }
                }
            }

            return baseSheetIndices;
        }


        private void DeleteUnmatchedRows(Excel.Worksheet sheet, List<int> rowIndices, bool headerRow)
        {
            var totalNumOfRows = sheet.UsedRange.Rows.Count;
            int deletedRows = 0; // This is needed to prevent deleting the incorrect rows

            for (int i = (!headerRow) ? 1 : 2; i <= totalNumOfRows; i++)
            {
                if (!rowIndices.Contains(i))
                {
                    Range row = sheet.Rows[i - deletedRows];
                    row.Delete();
                    deletedRows++;
                }
            }
        }


        private static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


    }
}
