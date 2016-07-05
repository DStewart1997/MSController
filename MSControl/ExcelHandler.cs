﻿using System;
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
        static Excel.Workbooks workbooks = null;
        static Excel.Workbook workbook = null;
        static Excel.Sheets worksheets = null;
        static Excel.Worksheet worksheet = null;
        static Excel.Range range = null;
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

            if (sheet == "")
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

            range = worksheet.Range["A" + 1];
        }

        /// <summary>
        /// Creates an excel spreadsheet.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be created.</param>
        public void Create(string filePath)
        {
            excelApp = excelApp ?? new Excel.Application();
            if (excelApp == null)
                throw new Exception("Excel could not be started. Ensure it is correctly installed on the machine.");

            excelApp.Visible = false;
            workbooks  = workbooks  ?? excelApp.Workbooks;
            workbook   = workbook   ?? workbooks.Add(missing);
            worksheets = worksheets ?? workbook.Worksheets;
            worksheet  = worksheet  ?? (Excel.Worksheet)workbook.ActiveSheet;
            range      = range      ?? worksheet.Range["A" + 1];
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
            if (range != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            if (worksheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            if (worksheets != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheets);
            if (workbook != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            if (workbooks != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            if (excelApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

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
        public void RenameSheet(string newSheet, string oldSheet = "")
        {
            if (oldSheet != "")
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
            range = worksheet.Range[column + row];
            string cellValue = range.Value.ToString();

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
            int counter = 1;
            range = worksheet.Range[column + counter];
            string lastCell = "";

            while (range.Value != null)
            {
                lastCell = range.Value.ToString();
                counter++;
                range = worksheet.Range[column + counter];
            }

            return lastCell;
        }

        /// <summary>
        /// Gets a list of each value in the specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">The column to search.</param>
        /// <returns></returns>
        public List<string> GetAllInColumn(string column)
        {
            List<string> columnData = new List<string>();
            int counter = 1;
            range = worksheet.Range[column + counter];
            string lastCell = "";

            while (range.Value != null)
            {
                lastCell = range.Value.ToString();
                counter++;
                range = worksheet.Range[column + counter];
                columnData.Add(lastCell);
            }

            return columnData;
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
            List<string> columns = GetColumnList();

            int counter = 0;
            range = worksheet.Range[columns[counter] + row];
            string lastCell = "";

            while (range.Value != null)
            {
                lastCell = range.Value.ToString();
                counter++;
                range = worksheet.Range[columns[counter] + row];
            }

            return lastCell;
        }

        /// <summary>
        /// Gets a list of each value in the specified row in the open spreadsheet.
        /// </summary>
        /// <param name="row">The row to search.</param>
        /// <returns></returns>
        public List<string> GetAllInRow(int row)
        {
            List<string> columns = GetColumnList();
            List<string> rowData = new List<string>();

            int counter = 0;
            range = worksheet.Range[columns[counter] + row];
            string lastCell = "";

            while (range.Value != null)
            {
                lastCell = range.Value.ToString();
                counter++;
                range = worksheet.Range[columns[counter] + row];
                rowData.Add(lastCell);
            }

            return rowData;
        }


        // Write
        /// <summary>
        /// Writes a value to a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void WriteCell(string column, int row, string data, bool numberFormat = false)
        {
            range = worksheet.Range[column + row.ToString()];
            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";
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
            range = worksheet.Range[column + counter];

            while (range.Value != null)
            {
                counter++;
                range = worksheet.Range[column + counter];
            }

            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";
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
        public void DeleteColumn(string column)
        {
            // TODO
            throw new NotImplementedException();
        }

        /// <summary>
        /// Deletes the specified worksheet from the spreadsheet. If no sheet is specified the currently selected sheet is deleted.
        /// </summary>
        /// <param name="sheet">The sheet to delete.</param>
        public void DeleteSheet(string sheet = "")
        {
            if (sheet == "")
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
        private List<string> GetColumnList()
        {
            // Create a list for the columns from A-ZZZ
            List<string> columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Select(x => x.ToString()).ToList();  // A-Z

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
