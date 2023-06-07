using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using MetroFramework;
using Syncfusion.DocIO.DLS;

namespace MailMergeFileParseer
{
    public partial class Home : MetroFramework.Forms.MetroForm
    {
        string FileName,dt = "";
        private BackgroundWorker bw = new BackgroundWorker();
        public Home()
        {
            InitializeComponent();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.DoWork += new DoWorkEventHandler( DoWork);
           
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i <= 100; i++)
            {
                // Simulate long running work
                Thread.Sleep(100);
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            metroProgressBar1.Value = e.ProgressPercentage;
        }


        private void Home_Load(object sender, EventArgs e)
        {
            resetForm();
        }

        private void resetForm()
        {
            TmplateFileLabel.Text = "";
            TmplateFileLabel.Visible = false;
            ExcelFileLabel.Visible = false;
            ExcelFileLabel.Text = "";
            Result.Visible = false;
            metroProgressBar1.Hide();
        }

        private void ExcelFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse",

                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                DefaultExt = "xlsx",
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelFileLabel.Text = openFileDialog1.FileName;
                ExcelFileLabel.Visible = true;
            }
        }

        private void TemplateFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse",
                Filter = "docx files (*.docx)|*.docx|All files (*.*)|*.*",
                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "docx",
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                TmplateFileLabel.Text = openFileDialog1.FileName;
                TmplateFileLabel.Visible = true;
            }
        }

        private void Transform_Click(object sender, EventArgs e)
        {
            try
            {
                bool FilesSelected = true;
                Transform.Enabled = false;
                Result.Visible = false;
                if (string.IsNullOrEmpty(TmplateFileLabel.Text))
                {
                    Result.Text = "Template File Not Selected";
                    Result.UseStyleColors = true;
                    Result.Visible = true;
                    Result.ForeColor = System.Drawing.Color.Red;
                    FilesSelected = false;
                }
                if (string.IsNullOrEmpty(ExcelFileLabel.Text))
                {
                    Result.Text = "Excel File Not Selected";
                    Result.UseStyleColors = true;
                    Result.Visible = true;
                    Result.ForeColor = System.Drawing.Color.Red;
                    FilesSelected = false;
                }

                if (FilesSelected)
                {
                    FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                    folderBrowserDialog.Description = "Select Output Folder";
                    string OutputPath = "";
                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        OutputPath = folderBrowserDialog.SelectedPath;
                    }
                    
                    DataSet resultTable = null;
                    using (var stream = File.Open(ExcelFileLabel.Text, FileMode.Open, FileAccess.Read))
                    {
                        // Auto-detect format, supports:
                        //  - Binary Excel files (2.0-2003 format; *.xls)
                        //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            resultTable = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                        }
                    }
                    

                    System.Data.DataTable dataTable = new System.Data.DataTable("HTML");
                    DataRow datarow = dataTable.NewRow();
                    int counter = 0;
                    bw.RunWorkerAsync();
                    metroProgressBar1.Show();
                    dt = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    foreach (DataRow dr in resultTable.Tables[0].Rows)
                    {
                        counter++;
                        for (int i = 0; i < resultTable.Tables[0].Columns.Count; i++)
                        {
                            string columnName = resultTable.Tables[0].Columns[i].ColumnName;
                            string value = dr[resultTable.Tables[0].Columns[i].ColumnName].ToString();
                            switch (columnName)
                            {
                                case "Document":
                                    FileName = value;
                                    if (!dataTable.Columns.Contains(columnName))
                                        dataTable.Columns.Add(columnName);
                                    if (columnName.ToLower().Contains("html"))
                                    {
                                        datarow[columnName] = value.Replace("<br>", "<br/>");
                                    }
                                    else
                                        datarow[columnName] = value;
                                    break;
                                default:
                                    if (!dataTable.Columns.Contains(columnName))
                                        dataTable.Columns.Add(columnName);
                                    if (columnName.ToLower().Contains("html"))
                                    {
                                        datarow[columnName] = value.Replace("<br>", "<br/>");
                                    }
                                    else
                                        datarow[columnName] = value;
                                    break;
                            }

                        }
                        dataTable.Rows.Add(datarow);
                        using (WordDocument document = new WordDocument(Path.GetFullPath(TmplateFileLabel.Text)))
                        {
                            //Creates mail merge events handler to replace merge field with HTML
                            document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
                            //Gets data to perform mail merge

                            //Performs the mail merge
                            document.MailMerge.Execute(dataTable);
                            //Removes mail merge events handler
                            document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);
                            //Saves the Word document instance
                            document.Save(OutputPath + "\\" + FileName + ".docx");
                        }
                        using (WordprocessingDocument doc =
                        WordprocessingDocument.Open(OutputPath + "\\" + FileName + ".docx", true))
                        {
                            // Get a reference to the main document part.
                            var docPart = doc.MainDocumentPart;

                            // Count the header and footer parts and continue if there 
                            // are any.

                            if (docPart.HeaderParts.Count() > 0 ||
                                docPart.FooterParts.Count() > 0)
                            {
                                docPart.DeleteParts(docPart.HeaderParts);
                                docPart.DeleteParts(docPart.FooterParts);
                            }
                            Document document = docPart.Document;
                            var headers =
                         document.Descendants<HeaderReference>().ToList();
                            foreach (var header in headers)
                            {
                                header.Remove();
                            }




                            document.Save();
                            // Code removed here...
                        }

                        dataTable.Rows.Clear();
                        int percent = counter / resultTable.Tables[0].Rows.Count * 100;
                        bw.ReportProgress(percent);
                    }
                    bw.CancelAsync();
                    metroProgressBar1.Hide();
                    //metroProgressBar1.Value = 100;
                    Result.Text = "Mail Merged Files are created at:" + OutputPath;
                    Result.ForeColor = System.Drawing.Color.Green;
                    Result.Visible = true;
                }
            }
            catch(Exception ex)
            {
                Result.Text = ex.Message;
                Result.Visible = true;
                Result.ForeColor = System.Drawing.Color.Red;
            }
            Transform.Enabled = true;
        }

        #region Helper methods
        /// <summary>
        /// Replaces merge field with HTML string by using MergeFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.TableName.Equals("HTML"))
            {
                if (args.FieldName.ToLower().Contains("html"))
                {
                    try
                    {
                        string text = args.FieldValue as string;
                        string Html = "";
                        if (args.FieldName.ToLower() == "htmlcode" || args.FieldName.ToLower() == "htmlcontent")
                        {
                           
                            if (args.FieldValue.ToString().Contains("<body>"))
                            {
                                Html = args.FieldValue.ToString().Replace("<body>", "<body style=\"font-size: 5pt; font-family: Calibri Light\">");
                            }
                            else
                            {
                                Html = "<body style=\"font-size: 5pt; font-family: Calibri Light\"> " + args.FieldValue.ToString() + " </body>";
                            }
                            
                        }

                        if (args.FieldName.ToLower().Contains("exam") || args.FieldName.ToLower().Contains("email"))
                        {

                            if (args.FieldValue.ToString().Contains("<body>"))
                            {
                                Html = args.FieldValue.ToString().Replace("<body>", "<body style=\"font-size: 10pt; font-family: Calibri Light\">");
                            }
                            else
                            {
                                Html = "<p style=\"font-size: 10pt;font-family: Calibri Light;\" >" + args.FieldValue.ToString() + "</p>";
                            }

                        }


                        WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                        int paraIndex = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                        int fieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                        if (args.FieldValue.ToString() == "0" || string.IsNullOrEmpty(args.FieldValue.ToString()))
                        {
                            Html = args.FieldValue.ToString();
                            args.CharacterFormat.FontName = "Calibri Light";
                            args.CharacterFormat.FontSize = 10;
                        }
                        else
                        {
                            //Appends HTML string at the specified position of the document body contents

                           
                            paragraph.OwnerTextBody.InsertXHTML(Html.ToString(), paraIndex, fieldIndex);
                            //Resets the field value
                            args.Text = string.Empty;
                        }
                        
                    }
                    catch(Exception ex)
                    {
                        
                        File.AppendAllText("./logs_"+ dt + ".txt", "Error in Document:" +  FileName + ", Field Name:" + args.FieldName + "\n");
                        File.AppendAllText("./logs_" + dt + ".txt", ex.Message + "\n");
                        File.AppendAllText("./logs_" + dt + ".txt", "Original Field Value:" + args.FieldValue.ToString() + "\n");
                        File.AppendAllText("./logs_" + dt + ".txt","========================================================\n");
                    }
                }

                if ((args.FieldName.ToLower().Contains("exam") ||  args.FieldName.ToLower().Contains("email") ) )
                {
                    args.CharacterFormat.FontName = "Calibri Light";
                    args.CharacterFormat.FontSize = 10;
                }
            }
        }
       

        #endregion
    }
}
