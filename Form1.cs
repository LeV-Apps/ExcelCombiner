﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.CodeDom;

//recources
//https://stackoverflow.com/questions/10419071/using-c-sharp-to-read-write-excel-files-xls-xlsx
//https://www.encodedna.com/windows-forms/read-an-excel-file-in-windows-forms-application-using-csharp-vbdotnet.htm

namespace ExcelCombiner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            // Show the Open File dialog. If the user clicks OK, load the
            // file that the user chose.
            if (uploadExcelDialog.ShowDialog() == DialogResult.OK)
            {
                //saves the file path
                string filePath = uploadExcelDialog.FileName;

                //check if the file is an excel file
                if (!filePath.EndsWith(".xlsx"))
                {
                    //shows an error to the user and returns
                    errorUpload.SetError(uploadButton, "unzulässiges Dateiformat");
                    Debug.Print("unsupported file format");
                    return;
                }
                //removes error message (if it exists)
                errorUpload.Clear();

                //create new workbook object
                var wbTemplate = new XLWorkbook();
                var ws = wbTemplate.AddWorksheet();

                ws.Cell("B3").Value = "Hallo Welt";

                //overwrite the existing file
                wbTemplate.SaveAs(filePath);
                Debug.Print("selected file changed");
            }
            else
            {
                errorUpload.SetError(uploadButton, "upload fehlgeschlagen");
            }
        }
        private void createSampleButton_Click(object sender, EventArgs e)
        {
            //let the user select the folder to save the sample to
            if (selectSampleFolderDialog.ShowDialog() == DialogResult.OK)
            {
                //save the folder path
                string folderPath = selectSampleFolderDialog.SelectedPath;
                Debug.Print(folderPath);
                //create new workbook object
                var wbTemplate = new XLWorkbook();
                var ws = wbTemplate.AddWorksheet();

                //add the header cells
                ws.Cell("A1").Value = "Konto-Nr";
                ws.Cell("B1").Value = "Bezeichnung";
                ws.Cell("C1").Value = "Saldo";

                //saves the data to an excel file
                //creates a new one or overwrites an existing one
                string sampleFilePath = folderPath + "\\SUSA_Vorlage.xlsx";
                wbTemplate.SaveAs(sampleFilePath);
                Debug.Print("sample created");
            }
        }
    }
}
