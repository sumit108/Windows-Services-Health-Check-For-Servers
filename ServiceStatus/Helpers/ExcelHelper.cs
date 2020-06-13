using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.IO;
using static System.Configuration.ConfigurationManager;
namespace ServiceStatus
{
    public static class ExcelHelper
    {
        public static string filePath = AppSettings["ExcelFilePath"];
        public static List<string> xlReadSpecificColoumn(string sheet, int rowStart, int rowEnd, int column)
        {
            var data = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var xlSheet = package.Workbook.Worksheets[sheet];
                // Loop through all columns for reading column data
                for (var start = rowStart; start <= rowEnd; start++)
                {
                    if (xlSheet.Cells[start, column].Value == null) break;
                    // Read the cell data
                    data.Add(xlSheet.Cells[start, column].Value.ToString());
                }
            }
            return data;
        }

        public static List<string> xlReadSpecificColoumn(string sheet, int column, bool skipHeader = true)
        {
            var data = new List<string>();
            int rowStart;
            int rowEnd;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var xlSheet = package.Workbook.Worksheets[sheet];
                rowStart = skipHeader ? (xlSheet.Dimension.Start.Row) + 1 : xlSheet.Dimension.Start.Row;
                rowEnd = xlSheet.Dimension.End.Row;
                // Loop through all columns for reading column data
                for (var start = rowStart; start <= rowEnd; start++)
                {
                    if (xlSheet.Cells[start, column].Value == null) continue;
                    // Read the cell data
                    data.Add(xlSheet.Cells[start, column].Value.ToString());
                }
            }
            return data;
        }
        /// <summary>
        /// Loading data to dictionary
        /// </summary>
        public static Dictionary<string, string> GetDataIntoDictionary(string SheetNo, int KeyColumn, int ValueColumn)
        {
            Dictionary<string, string> Data = new Dictionary<string, string>();
            int RowStart = xlReadFirstRow(SheetNo, KeyColumn);
            int RowEnd = xlnumberofrows(SheetNo);
            for (int i = ++RowStart; i <= RowEnd; i++)
            {
                if (xlread(SheetNo, i, KeyColumn) != null)

                    Data.Add(xlread(SheetNo, i, KeyColumn), xlread(SheetNo, i, ValueColumn));
                else

                    continue;
            }

            return Data;
        }

        //get first row of an excel
        public static int xlReadFirstRow(string sheet, int KeyColumn)
        {
            int RowStart = 1;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var xlSheet = package.Workbook.Worksheets[sheet];
                var iRowStart = xlSheet.Dimension.Start.Row;
                // Read data from a cell

                for (int i = 1; i > 0; i++)
                {
                    if (xlread(sheet, RowStart, KeyColumn) == null)
                        RowStart++;
                    else
                    {
                        i = -1;
                        break;
                    }
                }
            }
            return RowStart;
        }

        //Get last row of an excel
        public static int xlnumberofrows(string sheet)
        {
            int iRowCnt;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var xlSheet = package.Workbook.Worksheets[sheet];

                iRowCnt = xlSheet.Dimension.End.Row;

            }
            return iRowCnt;
        }

        public static string xlread(string sheet, int row, int column)
        {
            string data;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var xlSheet = package.Workbook.Worksheets[sheet];
                // Read data from a cell
                data = xlSheet.Cells[row, column].Value != null ? xlSheet.Cells[row, column].Value.ToString() : null;
            }
            return data;
        }

        // Reads Excel and return datatable 
        public static DataTable ExcelRead(string sheet, string excelFilePath = "")
        {
            if (excelFilePath == "")
                excelFilePath = AppSettings["TestDataPath"];

            ExcelPackage excel = new ExcelPackage(new FileInfo(excelFilePath));
            var xlSheet = excel.Workbook.Worksheets[sheet];

            DataTable dt = new DataTable();
            return dt = ToDataTable(xlSheet);
        }

        // Convert Excel Data to datatable
        public static DataTable ToDataTable(this ExcelWorksheet ws, bool hasHeaderRow = true)
        {
            var tbl = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                tbl.Columns.Add(hasHeaderRow ?
                    firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            var startRow = hasHeaderRow ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow) row[cell.Start.Column - 1] = cell.Text;
                tbl.Rows.Add(row);
            }
            return tbl;
        }
    }
}
