using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Excel, can read and write to cells and also check if a spreadsheet is open.
    /// </summary>
    public class ExcelHandler
    {
        static Excel.Application excelApp = null;
        static Excel.Workbooks workbooks  = null;
        static Excel.Workbook workbook    = null;
        static Excel.Sheets worksheets    = null;
        static Excel.Worksheet worksheet  = null;
        object missing = System.Reflection.Missing.Value;

        /// <summary>
        /// Instantiates the class.
        /// </summary>
        public ExcelHandler()
        {
        }

        // Open, Close, Create, IsOpen
        /// <summary>
        /// Opens an excel spreadsheet for processing.  If it does not exist an exception will be thrown.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be opened.</param>
        /// <param name="sheet">The worksheet to open. If it does not exist an exception is thrown.</param>
        public void Open(string filePath, string sheet = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified file: " + filePath + ", cannot be found.");

            excelApp = excelApp ?? new Excel.Application();
            if (excelApp == null)
                throw new Exception("Excel could not be started. Ensure it is correctly installed on the machine.");

            excelApp.Visible = false;
            workbooks  = workbooks  ?? excelApp.Workbooks;
            workbook   = workbook   ?? workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            worksheets = worksheets ?? workbook.Worksheets;

            if (sheet == null)
            {
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            }
            else
            {
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.get_Item(sheet);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new ArgumentException("The specified worksheet: " + sheet + ", cannot be found.");
                }
            }
        }

        /// <summary>
        /// Creates an excel spreadsheet.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be created.</param>
        /// <param name="numberOfDefaultSheets">The number of default sheets to create the spreadsheet with.</param>
        public void Create(string filePath, int numberOfDefaultSheets = 3)
        {
            excelApp = excelApp ?? new Excel.Application();
            if (excelApp == null)
                throw new Exception("Excel could not be started. Ensure it is correctly installed on the machine.");

            excelApp.Visible = false;
            excelApp.SheetsInNewWorkbook = numberOfDefaultSheets;
            workbooks  = workbooks  ?? excelApp.Workbooks;
            workbook   = workbook   ?? workbooks.Add(missing);
            worksheets = worksheets ?? workbook.Worksheets;
            worksheet  = worksheet  ?? (Excel.Worksheet)workbook.ActiveSheet;

            workbook.SaveAs(filePath);
            Close();
        }

        /// <summary>
        /// Closes the excel spreadsheet.
        /// </summary>
        /// <param name="save">Boolean value of whether or not to save the file.</param>
        public void Close(bool save = false)
        {
            if (save)
                workbook.Save();

            workbook.Close(0);
            excelApp.Quit();

            ReleaseCOMObjects();
        }

        /// <summary>
        /// Releases all used COM objects, useful for when try, catch, finally blocks to ensure all COM objects are released.
        /// </summary>
        public void ReleaseCOMObjects()
        {
            // Release the COM objects
            if (worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }
            if (worksheets != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheets);
                worksheets = null;
            }
            if (workbook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
            }
            if (workbooks != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                workbooks = null;
            }
            if (excelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }

            // The internet said do this twice and it works so it's here
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Checks whether a spreadsheet is open or not.
        /// </summary>
        /// <param name="filePath">String value of the column of the cell.</param>
        /// <returns>
        /// True if the file is open, false if not.
        /// </returns>
        public bool IsOpen(string filePath)
        {
            // TODO
            throw new NotImplementedException();
        }


        // Navigation
        /// <summary>
        /// Adds a new sheet to the currently open spreadsheet and switches to it.
        /// </summary>
        /// <param name="sheet">The worksheet to create.</param>
        public void AddSheet(string sheet)
        {
            worksheet = (Excel.Worksheet)worksheets.Add(worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            worksheet.Name = sheet;
        }

        /// <summary>
        /// Renames the currently selected worksheet or a specified one.
        /// </summary>
        /// <param name="newSheet">The new name of the worksheet.</param>
        /// <param name="oldSheet">The worksheet to rename.</param>
        public void RenameSheet(string newSheet, string oldSheet = null)
        {
            if (oldSheet != null)
            {
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.get_Item(oldSheet);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new ArgumentException("The specified worksheet: " + oldSheet + ", was not found.");
                }

            }
            else
            {
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            }

            worksheet.Name = newSheet;
        }

        /// <summary>
        /// Switches from the current worksheet to the specified one.
        /// </summary>
        /// <param name="sheet">The worksheet to switch to. If it is not found it will instead switch to the default.</param>
        public void ChangeSheet(string sheet)
        {
            try
            {
                worksheet = (Excel.Worksheet)worksheets.get_Item(sheet);
            }
            catch (Exception)
            {
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            }
        }

        /// <summary>
        /// Gets a list of the names of the sheets in the workbook.
        /// </summary>
        /// <returns>String list of worksheet names.</returns>
        public IEnumerable<string> GetSheets()
        {
            var sheets = new List<string>();
            foreach (Excel.Worksheet sheet in workbook.Sheets)
                sheets.Add(sheet.Name);

            return sheets;
        }


        // Read
        /// <summary>
        /// Gets the value from a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <returns>
        /// The value from the specified cell.
        /// </returns>
        public string GetCell(string column, int row)
        {
            Excel.Range range = worksheet.Range[column + row];
            string cellValue = range.Value.ToString();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
            return cellValue;
        }

        /// <summary>
        /// Gets the value from the last cell in a specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">The column to search.</param>
        /// <returns>
        /// The value from the last cell in the specified column.
        /// </returns>
        public string GetLastCellInColumn(string column)
        {
            int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;  // Last occupied row in entire spreadsheet
            string lastCell = "";
            Excel.Range range = null;

            // Go up through each row until a non null value is found, or we get to the top
            while (range?.Value == null && lastRow > 0)
            {
                range = worksheet.Range[column + lastRow];
                lastRow--;
            }

            lastCell = range.Value?.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
            return lastCell;
        }

        /// <summary>
        /// Gets a list of each value in the specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">The column to search.</param>
        /// <param name="includeBlanks">Whether to include blank spaces between first and last cells. Doesn't include trailing blank lines (ie. after last cell in column).</param>
        /// <returns>An enumerable list of each occupied cell in the specified column. Blanks are not included.</returns>
        public IEnumerable<string> GetAllInColumn(string column, bool includeBlanks = false)
        {
            int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;  // Last occupied row in entire spreadsheet
            IList<string> columnData = new List<string>();
            string lastCell = "";
            Excel.Range range = null;
            bool hitValid = false;

            while (lastRow > 0)
            {
                range = worksheet.Range[column + lastRow];

                if (range.Value != null)
                {
                    lastCell = range.Value.ToString();
                    columnData.Add(lastCell);
                    hitValid = true;
                }
                else if (includeBlanks && hitValid)
                {
                    columnData.Add("");
                }

                lastRow--;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
            return columnData.Reverse();  // Reverse return data as we start from the bottom
        }

        /// <summary>
        /// Gets the value from the last cell in a specified row in the open spreadsheet.
        /// </summary>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <returns>
        /// The value from the last cell in the specified row.
        /// </returns>
        public string GetLastCellInRow(int row)
        {
            int lastColumn = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column - 1;  // Last occupied column in entire spreadsheet
            IEnumerable<string> columns = GetColumnList();
            string lastCell = "";
            Excel.Range range = null;

            while (range?.Value == null && lastColumn > 0)
            {
                range = worksheet.Range[columns.ElementAt(lastColumn) + row];
                lastColumn--;
            }

            lastCell = range.Value?.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
            return lastCell;
        }

        /// <summary>
        /// Gets a list of each value in the specified row in the open spreadsheet.
        /// </summary>
        /// <param name="row">The row to search.</param>
        /// <param name="includeBlanks">Whether to include blank spaces between first and last cells. Doesn't include trailing blank lines (ie. after last cell in row).</param>
        /// <returns></returns>
        public IEnumerable<string> GetAllInRow(int row, bool includeBlanks = false)
        {
            int lastColumn = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column - 1;  // Last occupied column in entire spreadsheet
            IEnumerable<string> columns = GetColumnList();
            IList<string> rowData = new List<string>();
            string lastCell = "";
            Excel.Range range = null;
            bool hitValid = false;

            while (lastColumn > 0)
            {
                range = worksheet.Range[columns.ElementAt(lastColumn) + row];

                if (range.Value != null)
                {
                    lastCell = range.Value.ToString();
                    rowData.Add(lastCell);
                    hitValid = true;
                }
                else if (includeBlanks && hitValid)
                {
                    rowData.Add("");
                }

                lastColumn--;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
            return rowData.Reverse();  // Reverse return data as we start from the bottom
        }


        // Write - TODO fix blank line issue
        /// <summary>
        /// Writes a value to a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void WriteCell(string column, int row, string data, bool numberFormat = false)
        {
            Excel.Range range = worksheet.Range[column + row.ToString()];
            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
        }

        /// <summary>
        /// Writes a value to the last cell in a specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void WriteLastCellInColumn(string column, string data, bool numberFormat = false)
        {
            int counter = 1;
            Excel.Range range = worksheet.Range[column + counter];

            while (range.Value != null)
            {
                counter++;
                range = worksheet.Range[column + counter];
            }

            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
        }


        // Delete
        /// <summary>
        /// Deletes the specified row from the spreadsheet.
        /// </summary>
        /// <param name="row">The row to delete.</param>
        public void DeleteRow(int row)
        {
            // TODO

            throw new NotImplementedException();
        }

        /// <summary>
        /// Deletes the specified column from the spreadsheet.
        /// </summary>
        /// <param name="column">The column to delete.</param>
        /// <param name="shift">Whether to shift the columns to the left or just leave them blank.</param>
        public void DeleteColumn(string column, bool shift=true)
        {
            Excel.Range range = worksheet.Range[column + 1];

            if (shift)
                range.EntireColumn.Delete(missing);
            else
                range.EntireColumn.Clear();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
        }

        /// <summary>
        /// Deletes the specified worksheet from the spreadsheet. If no sheet is specified the currently selected sheet is deleted.
        /// </summary>
        /// <param name="sheet">The sheet to delete.</param>
        public void DeleteSheet(string sheet = null)
        {
            if (sheet == null)
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            else
            {
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.get_Item(sheet);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new ArgumentException("The specified worksheet: " + sheet + ", was not found.");
                }
            }

            worksheet.Delete();
        }


        // Misc
        private IEnumerable<string> GetColumnList()
        {
            // Create a list for the columns from A-ZZZ
            IList<string> columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Select(x => x.ToString()).ToList();  // A-Z

            for (int i = 0; i < 26; i++)
                for (int j = 0; j < 26; j++)
                    columns.Add(columns[i] + columns[j]);  // AA-ZZ

            for (int i = 0; i < 26; i++)
                for (int j = 0; j < 26; j++)
                    for (int k = 0; k < 26; ++k)
                        columns.Add(columns[i] + columns[j] + columns[k]);  // AAA-ZZZ

            return columns;
        }
    }

}
