using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using static ExcelCombiner.Form1;

namespace ExcelCombiner
{
    public class FileEditer
    {
        /// <summary>
        /// Throws all balances into one file, doesnt sort it or combines duplicates tho
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="uploadedFiles"></param>
        /// <returns></returns>
        public static XLWorkbook CombineFiles(XLWorkbook workbook, List<UploadedFile> uploadedFiles)
        {
            //starts combining the files
            IXLWorksheet comb_ws;
            if (workbook == null)
            {
                workbook = new XLWorkbook();
                comb_ws = workbook.AddWorksheet();
            }
            else
            {
                comb_ws = workbook.Worksheet(1);
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
            return workbook;
        }
        /// <summary>
        /// Checks which balances have the same number and description and adds them together
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public static XLWorkbook CombineDuplicates(XLWorkbook workbook,Label outputTextBox)
        {
            var worksheet = workbook.Worksheet(1);

            var firstRow = worksheet.FirstRowUsed();
            var lastRow = worksheet.LastRowUsed();

            var currentRow = firstRow;
            if (currentRow == null || currentRow == lastRow) return null;
            var rowToCompare = firstRow.RowBelow();

            while (true)
            {
                //default message if everything works flawless, is overwritten if thats not the case
                outputTextBox.Text = "Alle Zeilen konnten erfolgreich zusammengefügt werden, die " +
                    "Datei kann nun gedownloaded werden.";

                //compare the current row and the row to compare, there are 3 possible scenarios
                //1. they both have the same number and discription -> add them together
                //2. they have destinct numbers and descriptions -> dont modify them
                //3. they have either the same number OR description -> mark the cells and tell the user

                if (currentRow.Cell("A").Value.ToString() == rowToCompare.Cell("A").Value.ToString() &&
                    currentRow.Cell("B").Value.ToString() == rowToCompare.Cell("B").Value.ToString())
                {
                    Debug.Print("rows are identical");
                    Debug.Print(currentRow.Cell("A").Value.ToString() + " = " + rowToCompare.Cell("A").Value.ToString());
                    Debug.Print(currentRow.Cell("B").Value.ToString() + " = " + rowToCompare.Cell("B").Value.ToString());
                    //Scenario 1, they have the same number and description -> add together
                    double balanceCurrentRow = currentRow.Cell("C").GetDouble();
                    double balanceRowToCompare = rowToCompare.Cell("C").GetDouble();
                    double newBalance = balanceCurrentRow + balanceRowToCompare;
                    Debug.Print("new value = " + newBalance + " old values : " + balanceCurrentRow.ToString() +
                                " " + balanceRowToCompare.ToString());
                    currentRow.Cell("C").Value = newBalance;

                    //remove the now added row
                    if (rowToCompare == lastRow)
                    {
                        rowToCompare.Delete();
                        lastRow = worksheet.LastRowUsed();
                        if (currentRow == lastRow) break;
                        currentRow = currentRow.RowBelow();
                    } else
                    {
                        rowToCompare.Delete();
                    }
                    if (currentRow == lastRow) break;
                    rowToCompare = currentRow.RowBelow();

                    //if (rowToCompare == null) break;

                } else if (currentRow.Cell("A").Value.ToString() != rowToCompare.Cell("A").Value.ToString() &&
                           currentRow.Cell("B").Value.ToString() != rowToCompare.Cell("B").Value.ToString())
                {
                    //Scenario 2, they are completely destinct, go on with the script
                    Debug.Print("rows are destinct");
                    Debug.Print(currentRow.Cell("A").Value.ToString() + " = " + rowToCompare.Cell("A").Value.ToString());
                    Debug.Print(currentRow.Cell("B").Value.ToString() + " = " + rowToCompare.Cell("B").Value.ToString());

                    if (rowToCompare != lastRow)
                    {
                        rowToCompare = rowToCompare.RowBelow();
                    } else {
                        //currentRow was compared with all other rows, go on with next row
                        currentRow = currentRow.RowBelow();
                        if (currentRow == lastRow || currentRow == null) break;
                        rowToCompare = currentRow.RowBelow();
                    }
                }
                else
                {
                    //Scenario 3, they could be the same, mark the cell and tell user
                    Debug.Print("rows could be the same or destinct");
                    Debug.Print(currentRow.Cell("A").Value.ToString() + " = " + rowToCompare.Cell("A").Value.ToString());
                    Debug.Print(currentRow.Cell("B").Value.ToString() + " = " + rowToCompare.Cell("B").Value.ToString());
                    //mark rows
                    currentRow.Style.Fill.BackgroundColor = XLColor.Red;
                    rowToCompare.Style.Fill.BackgroundColor = XLColor.Red;
                    //tell user, overwrites the default message
                    outputTextBox.Text = "Manche Zeilen konnten nicht zusammengefügt werden, da dort lediglich die " +
                        "Kontonummer ODER lediglich der Kontennahme übereinstimmen. Diese müssen manuell bearbeitet werden, " +
                        "die jeweiligen Zeilen wurden rott markiert. Die Datei kann nun gedownloaded werden";
                    //go on with the script
                    if (rowToCompare != lastRow)
                    {
                        rowToCompare = rowToCompare.RowBelow();
                    }
                    else
                    {
                        //currentRow was compared with all other rows, go on with next row
                        currentRow = currentRow.RowBelow();
                        if (currentRow == lastRow || currentRow == null) break;
                        rowToCompare = currentRow.RowBelow();
                    }
                }
            }
            
            return workbook;
        }
    }
}
