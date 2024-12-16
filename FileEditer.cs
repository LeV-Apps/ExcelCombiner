using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using static ExcelCombiner.Form1;

namespace ExcelCombiner
{
    public class FileEditer
    {
        /// <summary>
        /// Throws all balances into one file, doesnt sort it or combines duplicates tho
        /// </summary>
        /// <param name="combined_wb"></param>
        /// <param name="uploadedFiles"></param>
        /// <returns></returns>
        public static XLWorkbook CombineFiles(XLWorkbook combined_wb, List<UploadedFile> uploadedFiles)
        {
            //starts combining the files
            IXLWorksheet comb_ws;
            if (combined_wb.Worksheets.Count() == 0)
            {
                comb_ws = combined_wb.AddWorksheet();
            }
            else
            {
                comb_ws = combined_wb.Worksheet(1);
            }
            int rowIndex = 1;

            foreach (var uploadedFile in uploadedFiles)
            {
                //gets the first worksheet and add it to the combined worksheet
                var worksheet = uploadedFile.workbook.Worksheet(1);
                var firstRow = worksheet.FirstRowUsed();
                var lastRow = worksheet.LastRowUsed();
                var currentRow = firstRow;

                while (true)
                {
                    var combRow = comb_ws.Row(rowIndex);
                    combRow.Cell("A").Value = currentRow.Cell("A").Value;
                    combRow.Cell("B").Value = currentRow.Cell("B").Value;
                    combRow.Cell("C").Value = currentRow.Cell("C").Value;

                    rowIndex++;

                    if (currentRow == lastRow)
                    {
                        break;
                    }

                    currentRow = currentRow.RowBelow();
                }
            }
            return combined_wb;
        }
    }
}
