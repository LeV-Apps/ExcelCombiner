using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelCombiner
{
    public class FileValidation
    {
        /// <summary>
        /// Check if the worksheet has the desired formatting (like the sample)
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static IXLWorksheet CheckFileFormat(IXLWorksheet ws)
        {
            var wb = new XLWorkbook();
            var cleanws = wb.AddWorksheet(); 

            //1. find the first row that isnt the header
            var row = ws.FirstRowUsed();
            var lastRow = ws.LastRowUsed();
            int validRows = 0;

            if (row == null)
            {
                return null;
            }

            while (true)
            {
                if (CheckFirstCell(row))
                {
                    validRows++;
                    var cleanRow = cleanws.Row(validRows);
                    cleanRow.Cell("A").Value = row.Cell("A").Value;
                    cleanRow.Cell("B").Value = row.Cell("B").Value;
                    cleanRow.Cell("C").Value = row.Cell("C").Value;
                    cleanRow = row;
                }
                //is it the last row ?
                if (row == lastRow)
                {
                    //last row reached
                    break;
                }
                //might be still the header ?, check next row
                var oldRow = row;
                row = row.RowBelow();

                //delete row
                //oldRow.Delete();
            }
            if (validRows > 0)
            {
                return cleanws;
            }
            return null;
        }
        private static bool CheckFirstCell(IXLRow row)
        {
            var firstCell = row.Cell("A");
            if (firstCell.DataType == XLDataType.Number)
            {
                //everything is ok, now check the second cell.
                if (CheckSecondCell(row)) return true;
            }
            else if (firstCell.DataType == XLDataType.Text)
            {
                //check if the string is numeric
                if (int.TryParse(firstCell.GetString(), out int firstCellValue))
                {
                    //everything is ok, now check the second cell.
                    if (CheckSecondCell(row)) return true;
                }
            }
            return false;
        }
        private static bool CheckSecondCell(IXLRow row)
        {
            //the second cell is the balance description -> string
            var secondCell = row.Cell("B");
            if (secondCell.DataType == XLDataType.Text)
            {
                //second cell also has the desired format, check last cell
                if (CheckThirdCell(row)) return true;
            }
            return false;
        }
        private static bool CheckThirdCell(IXLRow row)
        {
            //last cell contails the value as an currency, must be an float/int
            var thirdCell = row.Cell("C");
            if (thirdCell.DataType == XLDataType.Number)
            {
                //all cells have the desired format, continue with the script
                return true;
            }
            return false;
        }
    }
}
