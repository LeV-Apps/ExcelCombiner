using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Diagnostics;
using static ExcelCombiner.Form1;


namespace ExcelCombiner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class UploadedFile
        {
            public Label label = new Label();
            public string filePath;
            public string fileName;
            public Button removeButton = new Button();
            public XLWorkbook workbook = new XLWorkbook();

        }

        int maxFileCount = 8;
        int maxFilenameDisplayLength = 50;
        List<UploadedFile> uploadedFiles = new List<UploadedFile>();

        XLWorkbook combined_wb = null;

        private void uploadButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (uploadedFiles.Count >= maxFileCount)
                {
                    errorField.Clear();
                    errorField.SetError(uploadButton, "maximale Anzahl an Dateien erreicht");
                    outputConsole.Text = "Die maximale anzahl der hochladbaren Dateien ist erreicht, " +
                        "falls dennoch mehr Dateien bearbeitet werden müssen kann das Programm " +
                        "auch mehrere male durchlaufen werden.";
                    return;
                }
                // Show the Open File dialog. If the user clicks OK, load the
                // file that the user chose.
                if (uploadFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //saves the file path
                    string filePath = uploadFileDialog.FileName;

                    //check if the file is an excel file
                    if (!filePath.EndsWith(".xlsx"))
                    {
                        //shows an error to the user and returns
                        errorField.Clear();
                        errorField.SetError(uploadButton, "unzulässiges Dateiformat");
                        outputConsole.Text = "Upload ist fehlgeschlagen, handelt es sich bei " +
                            "der ausgesuchten Datei um eine Excel Datei ? (.xlsx)";
                        Debug.Print("unsupported file format");
                        return;
                    }

                    //check if the file contains the desired format
                    var wb = new XLWorkbook(filePath);
                    var ws = wb.Worksheet(1); //index starts at 1, not 0

                    //check if the file has the desired format
                    var cleanedWs = FileValidation.CheckFileFormat(ws);
                    if (cleanedWs == null)
                    {
                        errorField.Clear();
                        errorField.SetError(uploadButton, "ungültige Formatierung/Werte");
                        outputConsole.Text = "Upload ist fehlgeschlagen, da die Datei ein " +
                            "ungültiges Format hat. Siehe die Vorlage um zu sehen wie die Datei aufgebaut " +
                            "sein soll.";
                        return;
                    }
                    //ads now cleaned worksheet and remove the old one
                    wb.AddWorksheet(cleanedWs);
                    wb.Worksheet(1).Delete();

                    //removes error message (if it exists)
                    errorField.Clear();

                    //creates a new uploaded file element
                    UploadedFile uploadedFile = new UploadedFile();

                    //save the path of the selected file
                    uploadedFile.filePath = filePath;

                    //saves the cleaned up workbook
                    uploadedFile.workbook = wb;

                    //create an textfield with the name of the file
                    uploadedFile.label.Name = "file01Label";
                    char[] seperator = "\\".ToCharArray();
                    string[] filePathSplitted = filePath.Split(seperator);
                    string fileName = filePathSplitted.Last();
                    if (fileName.Length > maxFilenameDisplayLength)
                    {
                        uploadedFile.label.Text = fileName.Substring(0, maxFilenameDisplayLength);
                    } else
                    {
                        uploadedFile.label.Text = fileName;
                    }
                    uploadedFile.fileName = fileName;

                    //define its size and dock it in the field right of the button
                    uploadedFile.label.Size = new Size(uploadedFile.label.PreferredWidth,
                                                      uploadedFile.label.PreferredHeight);
                    uploadedFile.label.Parent = flowLayoutPanel2;
                    uploadedFile.label.TextAlign = ContentAlignment.MiddleCenter;

                    //create an button to remove the selected file
                    uploadedFile.removeButton.Name = "fileRemoveButton";
                    uploadedFile.removeButton.Size = uploadedFile.removeButton.PreferredSize;
                    uploadedFile.removeButton.Text = "Datei entfernen";
                    uploadedFile.removeButton.AutoSize = true;
                    uploadedFile.removeButton.Parent = flowLayoutPanel2;

                    //Adds an listener to the button which is called when the button is pressed
                    uploadedFile.removeButton.Click += new EventHandler(delegate (Object o, EventArgs a)
                    {
                        //remove the uploaded file
                        uploadedFile.label.Dispose();
                        uploadedFile.removeButton.Dispose();
                        uploadedFiles.Remove(uploadedFile);
                        //removes error message (if it exists)
                        errorField.Clear();
                        outputConsole.Text = "Datei : " +uploadedFile.fileName +" wurde entfernt";
                    });
                    uploadedFiles.Add(uploadedFile);
                    outputConsole.Text = "Upload von : " +uploadedFile.fileName+ " war erfolgreich";

                }
                else
                {
                    errorField.Clear();
                    errorField.SetError(uploadButton, "upload fehlgeschlagen");
                    outputConsole.Text = "Upload ist fehlgeschlagen, Problem war warscheinlich das " +
                        "keine Datei ausgewählt wurde";
                }
            }
            catch (Exception error)
            {
                PrintError(error, errorField, "unbekannter Fehler beim Upload");
            }
        }

        

        private void createSampleButton_Click(object sender, EventArgs e)
        {
            try
            {
                //let the user select the folder to save the sample to
                if (selectFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    errorField.Clear();

                    //save the folder path
                    string folderPath = selectFolderDialog.SelectedPath;
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
                    outputConsole.Text = "Vorlagen Datei wurde erfolgreich unter " +
                        sampleFilePath + " erstellt";
                    Debug.Print("sample created");
                } else
                {
                    errorField.Clear();
                    errorField.SetError(createSampleButton, "kein Speicherort ausgewählt");
                    outputConsole.Text = "Donwload der Vorlage abgebrochen, es wurde " +
                        "kein Speicherort ausgewählt";

                }
            }
            catch (Exception error)
            {
                PrintError(error, errorField, "unbekannter Fehler bei der Vorlagenerstellung");
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (uploadedFiles.Count < 1)
                {
                    errorField.Clear();
                    errorField.SetError(startButton, "es sind noch keine Dateien hochgeladen");
                    outputConsole.Text = "Programm kann noch nicht ausgeführt werden, da noch " +
                        "keine Datein hochgeladen wurden";
                    return;
                }
                errorField.Clear();

                //puts the content of all files into one file
                combined_wb = FileEditer.CombineFiles(combined_wb,uploadedFiles);

                //sorts the newly created file
                var sorted_wb = combined_wb;
                var ws = sorted_wb.Worksheet(1);
                var sorted_ws = ws.Sort(1);
                combined_wb = sorted_wb;

                //now combine the duplicates, if not possible marks them and tells the user
                combined_wb = FileEditer.CombineDuplicates(combined_wb,outputConsole);
            }
            catch (Exception error)
            {
                PrintError(error,null,"unbekannter fehler beim Zusammenfügen");
            }
        }

        private void downloadButton_Click(object sender, EventArgs e)
        {
            try
            {
                //check if there already an file to download
                if (combined_wb == null)
                {
                    errorField.Clear();
                    errorField.SetError(downloadButton, "es gibt noch keine Datei zum download");
                    outputConsole.Text = "Download ist fehlgeschlagen, da noch keine Datei zum " +
                        "Download bereit steht. Wurde das Programm überhaupt schon ausgeführt ?";
                    return;
                }
                errorField.Clear();

                //let the user select the folder to save the file to
                if (selectFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    //save the folder path
                    string folderPath = selectFolderDialog.SelectedPath;
                    Debug.Print(folderPath);

                    string newFilePath = folderPath + "\\" + "kombinierteSUSA.xlsx";
                    combined_wb.SaveAs(newFilePath);
                    outputConsole.Text = "Datei wurden erfolgreich gedownloaded, Speicherort : " + folderPath;
                }
            }
            catch (Exception error)
            {
                PrintError(error, errorField, "unbekannter Fehler beim download");
            }
        }
        private void PrintError(Exception error,ErrorProvider errorProvider, string consoleMessage)
        {
            errorProvider.Clear();

            Debug.Print(error.Message);
            Debug.Print(error.StackTrace);

            string errorMessage = error.Message + " " + error.StackTrace;
            if (errorProvider != null)
            {
                errorProvider.SetError(uploadButton, combined_wb + " " + errorMessage);
            }
            outputConsole.Text = consoleMessage + " " +
                "Error Nachricht: " + error.Message + " Error ursprung:" +
                " " + error.StackTrace;
        }
    }
}
