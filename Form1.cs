using System;
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
using ClosedXML.Excel;
using System.Diagnostics;

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
            /*if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
            }*/
            //Workbook wb = Workbook.

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void createSampleButton_Click(object sender, EventArgs e)
        {
            //create 'workbook' object
            var wbTemplate = new XLWorkbook();
            var ws = wbTemplate.AddWorksheet();

            ws.Cell("B1").Value = "{{Model.Name}}";
            ws.Cell("B2").Value = "Children:";
            ws.Cell("B3").Value = "{{item.ChildName}}";

            ws.Cell("D2").Value = "Items in container:";
            ws.Cell("E2").Value = "{{item.ChildName}}";

            //saves the data to an existing excel file
            wbTemplate.SaveAs("C:/Users/mt/Documents/test2.xlsx");
            Debug.Print("Done");
            
        }
    }
}
