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
        HelpDialog helpDialog = null;
        private static int differencesSheetsCounter = 0;


        public Excel.Application ExcelApp
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }



        //--------------------------------------------------------------------------
        // Ribbon Modifiers
        //--------------------------------------------------------------------------
        #region

        public static void RepopulateSheetDropDowns()
        {
            var sheet1DropDown = Globals.Ribbons.Ribbon1.sheet1DropDown;
            var sheet2DropDown = Globals.Ribbons.Ribbon1.sheet2DropDown;

            sheet1DropDown.Items.Clear();
            sheet2DropDown.Items.Clear();

            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                RibbonDropDownItem item1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem(), item2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item1.Label = sheet.Name;
                item2.Label = sheet.Name;

                sheet1DropDown.Items.Add(item1);
                sheet2DropDown.Items.Add(item2);
            }

            sheet1DropDown.SelectedItemIndex = 0;
            sheet2DropDown.SelectedItemIndex = (sheet2DropDown.Items.Count > 1) ? 1 : -1;
        }



        #endregion



        //--------------------------------------------------------------------------
        // Event Handlers
        //--------------------------------------------------------------------------
        #region

        /// <summary>
        /// Refreshes the list of sheets in the drop downs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet1DropDown_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (e.Control.Id == "refreshButton1")
                RepopulateSheetDropDowns();
        }

        private void sheet2DropDown_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (e.Control.Id == "refreshButton2")
                RepopulateSheetDropDowns();
        }


        /// <summary>
        /// Removes all extraneous white space and commas from the textbox.
        /// i.e. a  ,b,,c = a,b,c
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Sheet1Range_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string sheet1Columns = (sender as RibbonEditBox).Text;

            // Remove all white space and extra commas from the edit box text
            sheet1Columns = sheet1Columns.Replace(" ", "");
            sheet1Columns = Regex.Replace(sheet1Columns, ",+", ",").Trim(',');

            (sender as RibbonEditBox).Text = sheet1Columns;

            string sheet1Name = sheet1DropDown.SelectedItem.Label;
            Excel.Worksheet sheet1 = null;

            foreach (Excel.Worksheet sheet in ExcelApp.Worksheets)
            {
                if (sheet.Name == sheet1Name)
                {
                    sheet1 = sheet;
                    break;
                }
            }

            SelectColumnsInRange(sheet1, (sender as RibbonEditBox).Text);
        }

        private void Sheet2Range_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string sheet2Columns = (sender as RibbonEditBox).Text;

            // Remove all white space and extra commas from the edit box text
            sheet2Columns = sheet2Columns.Replace(" ", "");
            sheet2Columns = Regex.Replace(sheet2Columns, ",+", ",").Trim(',');

            (sender as RibbonEditBox).Text = sheet2Columns;

            string sheet2Name = sheet1DropDown.SelectedItem.Label;
            Excel.Worksheet sheet2 = null;

            foreach (Excel.Worksheet sheet in ExcelApp.Worksheets)
            {
                if (sheet.Name == sheet2Name)
                {
                    sheet2 = sheet;
                    break;
                }
            }
            SelectColumnsInRange(sheet2, (sender as RibbonEditBox).Text);
        }

        /// <summary>
        /// Changes the button's label to indicate if it is checked or not
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet1HeaderToggle_Click(object sender, RibbonControlEventArgs e)
        {
            sheet1HeaderToggle.Label = sheet1HeaderToggle.Checked ? "Contains Header Row (✔)" : "Contains Header Row (   )";
        }

        private void sheet2HeaderToggle_Click(object sender, RibbonControlEventArgs e)
        {
            sheet2HeaderToggle.Label = sheet2HeaderToggle.Checked ? "Contains Header Row (✔)" : "Contains Header Row (   )";
        }

        /// <summary>
        /// Display the help dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void helpButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (helpDialog == null || helpDialog.IsDisposed)
            {
                helpDialog = new HelpDialog();
            }

            helpDialog.Show();
            helpDialog.Activate();
        }



        private void CompareSheetsButton_Click(object sender, RibbonControlEventArgs e)
        {
            string sheet1Name = sheet1DropDown.SelectedItem.Label, sheet2Name = sheet2DropDown.SelectedItem.Label;
            string sheet1Columns = sheet1RangeBox.Text, sheet2Columns = sheet2RangeBox.Text;
            bool sheet1HeaderRow = sheet1HeaderToggle.Checked, sheet2HeaderRow = sheet2HeaderToggle.Checked;
            bool ignoreCaps = capsCheckBox.Checked, ignoreSpecialChars = specialCharsCheckBox.Checked;

            if (ValidateSheetSelection() && ValidateColumnInput(ExcelApp.Worksheets[1], sheet1Columns, true) && ValidateColumnInput(ExcelApp.Worksheets[2], sheet2Columns, true))
            {
                string[] sheet1ColumnArray = sheet1Columns.Split(','), sheet2ColumnArray = sheet2Columns.Split(',');
                Excel.Worksheet sheet1 = null, sheet2 = null;

                foreach (Excel.Worksheet sheet in ExcelApp.Worksheets)
                {
                    if (sheet.Name == sheet1Name)
                        sheet1 = sheet;
                    else if (sheet.Name == sheet2Name)
                        sheet2 = sheet;
                }

                if (sheet1 != null && sheet2 != null)
                {
                    List<int> sheet1Indices = CompareLists(sheet1, sheet2, sheet1ColumnArray, sheet2ColumnArray, sheet1HeaderRow, sheet2HeaderRow, ignoreCaps, ignoreSpecialChars);
                    List<int> sheet2Indices = CompareLists(sheet2, sheet1, sheet2ColumnArray, sheet1ColumnArray, sheet2HeaderRow, sheet1HeaderRow, ignoreCaps, ignoreSpecialChars);
                    DisplayResults(sheet1, sheet2, sheet1Indices, sheet2Indices, sheet1HeaderRow, sheet2HeaderRow);
                }
                else
                {
                    MessageBox.Show("Cannot find one or both of the selected sheets. Please refresh the sheet drop down lists.");
                }
            }
        }

        private void SelectColumnsInRange(Excel.Worksheet sheet, string columns)
        {
            if (ValidateColumnInput(sheet, columns, false))
            {
                // Make sure the provided sheet is selected
                sheet.Select();

                string[] columnArray = columns.ToUpper().Split(',');
                StringBuilder columnRangeString = new StringBuilder();

                foreach (string column in columnArray)
                {
                    // A blank column will only occur if extra commas were placed after the columns
                    if (column == "")
                        continue;

                    columnRangeString.Append(column + ":" + column + ",");
                }
                columnRangeString.Length = columnRangeString.Length - 1;

                Range range = sheet.get_Range(columnRangeString.ToString(), Type.Missing);
                range.EntireColumn.Select();
            }
        }


        #endregion



        //--------------------------------------------------------------------------
        // Validation / Input Checking
        //--------------------------------------------------------------------------
        #region

        private bool ValidateSheetSelection()
        {
            // Make sure the user has selected two sheets
            if (sheet1DropDown.SelectedItemIndex == -1 || sheet2DropDown.SelectedItemIndex == -1)
            {
                MessageBox.Show("Two sheets must be selected");
                return false;
            }

            // Make sure the user has not selected the same sheet in both drop downs
            if (sheet1DropDown.SelectedItemIndex == sheet2DropDown.SelectedItemIndex)
            {
                MessageBox.Show("The first and second sheets cannot be the same");
                return false;
            }

            return true;
        }

        private bool ValidateColumnInput(Excel.Worksheet sheet, string columns, bool displayErrorMessages)
        {
            // Verify that there is input at all
            if (columns == "")
            {
                if (displayErrorMessages)
                    MessageBox.Show("A range of at least one column must be specified for " + sheet.Name, "Empty Range");
                return false;
            }

            // Verify that the column inputs contain only characters
            if (Regex.IsMatch(columns, "[^a-z|A-Z|,]"))
            {
                if (displayErrorMessages)
                    MessageBox.Show("The column input for " + sheet.Name + " can only contain commas and letters that represent columns", "Incorrect Input");
                return false;
            }

            // Input is correct
            return true;
        }


        #endregion



        //--------------------------------------------------------------------------
        // Comparison Processes
        //--------------------------------------------------------------------------
        #region

        private List<int> CompareLists(Excel.Worksheet sheet1, Excel.Worksheet sheet2, string[] sheet1Columns, string[] sheet2Columns, bool sheet1HeaderRow, bool sheet2HeaderRow, bool ignoreCaps, bool ignoreSpecialChars)
        {
            List<int> sheetIndices = new List<int>();

            //Gets the range and columns in the worksheet that are used. Range will be used to loop, and col to keep data intact
            var sheet1NumOfCols = sheet1.UsedRange.Columns.Count;
            var sheet1NumOfRows = sheet1.UsedRange.Rows.Count;

            var sheet2NumOfCols = sheet2.UsedRange.Columns.Count;
            var sheet2NumOfRows = sheet2.UsedRange.Rows.Count;

            // Compare each row in sheet1 to every row in sheet2. If a match is found, save the current row's index
            for (int i = (!sheet1HeaderRow) ? 1 : 2; i <= sheet1NumOfRows; i++)
            {
                StringBuilder sheet1RowString = new StringBuilder();
                bool matchFound = false;

                // Create the comparison string for the current row in sheet 1
                foreach (string column in sheet1Columns)
                {
                    // Add each cells' contents to the string
                    Range cell = sheet1.get_Range(column + i.ToString()); //sheet1.Cells[i, j] as Range;

                    if (cell.Value != null)
                    {
                        string value = Convert.ToString(cell.Value);

                        if (ignoreSpecialChars)
                            value = Regex.Replace(value, "[^0-9a-zA-Z]", "");

                        sheet1RowString.Append(value + ",");
                    }
                        
                }                

                // Compare the row in the base sheet with every row in the compare sheet
                for (int k = (!sheet2HeaderRow) ? 1 : 2; k <= sheet2NumOfRows; k++)
                {
                    StringBuilder sheet2RowString = new StringBuilder();

                    // Create the comparison string for the row in sheet 2
                    foreach (string column in sheet2Columns)
                    {
                        // Add each cells' contents to the string
                        Range cell = sheet2.get_Range(column + k.ToString());

                        if (cell.Value != null)
                        {
                            string value = Convert.ToString(cell.Value);

                            if (ignoreSpecialChars)
                                value = Regex.Replace(value, "[^0-9a-zA-Z]", "");

                            sheet2RowString.Append(value + ",");
                        }
                    }

                    // Compare the two rows and see if they are the same
                    if (ignoreCaps)
                    {
                        if (sheet1RowString.ToString().ToLower() == sheet2RowString.ToString().ToLower())
                        {
                            matchFound = true;
                            break;
                        }
                    }
                    else
                    {
                        if (sheet1RowString.ToString() == sheet2RowString.ToString())
                        {
                            matchFound = true;
                            break;
                        }
                    }

                    
                }

                // If a similar entry was not found in the other list, add it to the list
                if (!matchFound)
                {
                    sheetIndices.Add(i);
                }
            }

            return sheetIndices;
        }

        private void DisplayResults(Excel.Worksheet sheet1, Excel.Worksheet sheet2, List<int> sheet1Indices, List<int> sheet2Indices, bool sheet1HeaderRow, bool sheet2HeaderRow)
        {
            // Create a new sheet and place it at the end of all other sheets. Get a reference to the new sheet
            Excel.Worksheet lastSheet = ExcelApp.Worksheets[ExcelApp.Worksheets.Count];
            ExcelApp.Worksheets.Add(Type.Missing, lastSheet);
            Excel.Worksheet resultsSheet = ExcelApp.Worksheets[ExcelApp.Worksheets.Count];


            bool anyDifferenceSheets = false;

            // Reset the differences sheet counter if there are no such sheets
            foreach (Excel.Worksheet sheet in ExcelApp.Worksheets)
            {
                if (sheet.Name.Contains("Differences"))
                {
                    anyDifferenceSheets = true;
                    break;
                }
            }

            if (!anyDifferenceSheets)
                differencesSheetsCounter = 0;

            resultsSheet.Name = (differencesSheetsCounter == 0) ? "Differences" : "Differences " + differencesSheetsCounter.ToString();
            differencesSheetsCounter++;

            int currResultsRowIndex = 1;

            // Get the number of columns used by the header. The larger count will determine the size of the sheet names column widths
            int sheet1ColumnsCount = sheet1.UsedRange.Columns.Count;
            int sheet2ColumnsCount = sheet2.UsedRange.Columns.Count;
            int chosenColumnCount = (sheet1ColumnsCount > sheet2ColumnsCount) ? sheet1ColumnsCount : sheet2ColumnsCount;

            // Display the first sheet's name on the first row
            Range sheet1Name = resultsSheet.get_Range("A1", GetColNameFromIndex(chosenColumnCount) + "1");
            sheet1Name.Merge();
            sheet1Name.Font.Bold = true;
            sheet1Name.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet1Name.Value = sheet1.Name + " (Rows not contained in " + sheet2.Name + ")";
            currResultsRowIndex++;

            // Copy the header over, if there is one
            if (sheet1HeaderRow)
            {
                Range header = sheet1.Rows[1];
                Range resultsHeader = resultsSheet.Rows[currResultsRowIndex];
                header.Copy(resultsHeader);
                currResultsRowIndex++;
            }

            // Copy all of the rows from the first sheet to the new sheet
            int sheet1RowsCount = sheet1.UsedRange.Rows.Count;

            for (int i = (!sheet1HeaderRow) ? 1 : 2; i <= sheet1RowsCount; i++)
            {
                if (sheet1Indices.Contains(i))
                {
                    Range sheet1Row = sheet1.Rows[i];
                    Range resultsRow = resultsSheet.Rows[currResultsRowIndex];
                    sheet1Row.Copy(resultsRow);
                    currResultsRowIndex++;
                }
            }
            currResultsRowIndex++;


            // Display the second sheet's name on the first row
            Range sheet2Name = resultsSheet.get_Range("A" + currResultsRowIndex.ToString(), GetColNameFromIndex(chosenColumnCount) + currResultsRowIndex.ToString());
            sheet2Name.Merge();
            sheet2Name.Font.Bold = true;
            sheet2Name.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet2Name.Value = sheet2.Name + " (Rows not contained in " + sheet1.Name + ")";
            currResultsRowIndex++;

            // Copy the header over, if there is one
            if (sheet2HeaderRow)
            {
                Range header = sheet2.Rows[1];
                Range resultsHeader = resultsSheet.Rows[currResultsRowIndex];
                header.Copy(resultsHeader);
                currResultsRowIndex++;
            }

            // Copy all of the rows from the second sheet to the new sheet
            int sheet2RowsCount = sheet2.UsedRange.Rows.Count;

            for (int i = (!sheet2HeaderRow) ? 1 : 2; i <= sheet2RowsCount; i++)
            {
                if (sheet2Indices.Contains(i))
                {
                    Range sheet2Row = sheet2.Rows[i];
                    Range resultsRow = resultsSheet.Rows[currResultsRowIndex];
                    sheet2Row.Copy(resultsRow);
                    currResultsRowIndex++;
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

        #endregion

    }
}
