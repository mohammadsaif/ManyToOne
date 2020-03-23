using Microsoft.Office.Core;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Security;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;
using System.Runtime.InteropServices;

namespace ManyToOne
{
    public partial class ManyToOneForm : Form
    {
        public DataTable dataTableMergeData = new DataTable();
        public string filePath = string.Empty;
        public string selectedDataSheet = string.Empty;

        public ManyToOneForm()
        {
            InitializeComponent();
            LoadOutlook();
            
        }

        void LoadOutlook()
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.Accounts accounts = app.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                try
                {
                    comboBoxSendFrom.Items.Add(account.SmtpAddress);
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
            if (comboBoxSendFrom.Items.Count > 0)
            {
                comboBoxSendFrom.SelectedIndex = 0;
            }
        }

        public DataTable ReadExcel(string fileName)
        {
            string conn = string.Empty;
            DataTable dtSheets = null;
            DataTable dtexcel = new DataTable();
            string fileExt = Path.GetExtension(fileName);
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HRD=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    dtSheets = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dtSheets == null)
                    {
                        return null;
                    }

                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + selectedDataSheet + "]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
            return dtexcel;
        }



        private DataTable OpenFile()
        {
            string fileExt = string.Empty;
            if (filePath != string.Empty)
            {
                fileExt = Path.GetExtension(filePath);

            }
            DataTable dataTable = new DataTable();
            try
            {
                if (filePath != string.Empty)
                {
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        //Read worksheet into data table
                        dataTable = ReadExcel(filePath);
                    }
                }
            }
            catch (SecurityException ex)
            {
                // The user lacks appropriate permissions to read files, discover paths, etc.
                MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                    "Error message: " + ex.Message + "\n\n" +
                    "Details (send to Support):\n\n" + ex.StackTrace
                );
            }
            catch (Exception ex)
            {
                // Could not load the image - probably related to Windows file system permissions.
                MessageBox.Show("Cannot display the image: " + filePath.Substring(filePath.LastIndexOf('\\'))
                    + ". You may not have permission to read the file, or " +
                    "it may be corrupt.\n\nReported error: " + ex.Message);
            }
            //labelPath.Text = openFileDialogDataSource.FileName;
            return (dataTable);

        }

        void WriteDocumentProperty(string propertyName, string propertyValue)
        {
            Office.DocumentProperties properties;
            properties = Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
            if (propertyName == "ChildFields" && ReadDocumentProperty(propertyName) !=null)
            {
                propertyValue = ReadDocumentProperty(propertyName) + "|" + propertyValue;
            }
            if (ReadDocumentProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }
            properties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue);
        }

        void DeleteDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            properties = Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
            if (ReadDocumentProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }
            
            
        }
        void DeleteDocumentProperty(string propertyName, string valueToDelete)
        {
            Office.DocumentProperties properties;
            properties = Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

            if (ReadDocumentProperty(propertyName) != null && propertyName == "ChildFields")
            {
                string childFields = ReadDocumentProperty(propertyName);
                childFields = childFields.Replace(valueToDelete, "");
                properties[propertyName].Delete();
                properties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, childFields);
            }


        }

        private string ReadDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            //dynamic properties = null;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties
;

            if (properties != null)
            {
                for (int i = 1; i <= properties.Count; i++)
                {
                    if (properties[i].Name == propertyName)
                    {
                        return properties[i].Value.ToString();
                    }

                }
            }

            return null;
        }

        private void LoadFields()
        {
            //string filePath = string.Empty;
            string mergeOutputTo = string.Empty;

            //Load Data source files path
            filePath = ReadDocumentProperty("DataSourcePath");
            if (filePath != null)
            {
                labelFilePath.Text = filePath;
            }
            else
            {
                labelFilePath.Text = "No Data Source is selected";
            }

            selectedDataSheet = ReadDocumentProperty("DataSourceSheet");
            

            //load field "merge outpu to"
            mergeOutputTo = ReadDocumentProperty("MergeOutputTo");
            if (mergeOutputTo != null)
            {
                comboBoxMergeOutputTo.SelectedItem = mergeOutputTo;
            }
            else
            {
                comboBoxMergeOutputTo.SelectedItem = "Select merge output";
            }

            if (filePath!=null && comboBoxMergeOutputTo.SelectedIndex!=0 && radioButtonManyToOne.Checked )
            {
                tabControlAdvancedMerge.TabPages[1].Enabled = true;
                btnNext.Enabled = true;
            }
            else
            {
                tabControlAdvancedMerge.TabPages[1].Enabled = false;
                btnNext.Enabled = false;
            }
        }

        private void LoadMergeData()
        {
            
            dataTableMergeData = OpenFile();
            string childFields = string.Empty;
            string[] childFieldList;
            char[] separotor = { '|' };

            if (dataTableMergeData != null)
            {
                listBoxFields.Items.Clear();
                comboBoxSubjectField.Items.Clear();
                comboBoxToField.Items.Clear();
                foreach (DataColumn dataColumn in dataTableMergeData.Columns)
                {
                    listBoxFields.Items.Add(dataColumn.ColumnName);
                    comboBoxToField.Items.Add(dataColumn.ColumnName);
                    comboBoxSubjectField.Items.Add(dataColumn.ColumnName);



                }
                if (ReadDocumentProperty("KeyField") != null)
                {
                    textBoxKeyField.Text = ReadDocumentProperty("KeyField");
                    listBoxFields.Items.Remove(ReadDocumentProperty("KeyField"));

                }
                if (ReadDocumentProperty("ChildFields") != null)
                {
                    childFields = ReadDocumentProperty("ChildFields");
                    childFieldList = childFields.Split(separotor, StringSplitOptions.RemoveEmptyEntries);
                    listBoxChildFields.Items.Clear();
                    foreach (string item in childFieldList)
                    {
                        listBoxChildFields.Items.Add(item);
                    }

                }

            }

        }

        private void btnStart_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormDataSource formDataSource = new FormDataSource();
            formDataSource.ShowDialog();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ManyToOneForm_Load(object sender, EventArgs e)
        {
            LoadFields();

        }

        private void ManyToOneForm_Activated(object sender, EventArgs e)
        {
            LoadFields();
        }

        private void comboBoxMergeOutputTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            WriteDocumentProperty("MergeOutputTo", comboBoxMergeOutputTo.SelectedItem.ToString());
            LoadFields();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (btnNext.Text == "Next")
            {
                LoadMergeData();
                tabControlAdvancedMerge.SelectedTab = tabControlAdvancedMerge.TabPages[1];
                btnNext.Text = "Merge";
            }
            else if (btnNext.Text == "Merge" & btnNext.Enabled)
            {
                Word.Fields wordFields = Globals.ThisAddIn.Application.ActiveDocument.Fields;
                if (wordFields.Count > 0 && dataTableMergeData != null)
                {
                    if (Globals.ThisAddIn.Application.ActiveDocument.Saved)
                    {
                        MessageBox.Show("File is saved");
                    }
                    else
                    {
                        Globals.ThisAddIn.Application.ActiveDocument.Save();
                    }
                    string keyField = ReadDocumentProperty("KeyField");
                    Outlook.Application outlookApp = new Outlook.Application();
                    Outlook.Accounts outlookAccounts = outlookApp.Session.Accounts;
                    Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                    Word.Document wordDocument = Globals.ThisAddIn.Application.ActiveDocument;
                    
                    var grouped = from table in dataTableMergeData.AsEnumerable()
                                group table by new { keyCol = table[textBoxKeyField.Text] } into grp
                                select new
                                {
                                    Value = grp.Key,
                                    ColumnValues = grp
                                };
                    foreach (var key in grouped)
                    {
                        Word.Application wordApplication = new Word.Application();
                        wordApplication.ShowAnimation = false;
                        wordApplication.Visible = false;
                        object missing = System.Reflection.Missing.Value;
                        Word.Document newDocument = wordApplication.Documents.Open(wordDocument.Path + @"\" + wordDocument.Name, ReadOnly:true);
                       
                        Word.MailMerge mailMerge = newDocument.MailMerge;


                        DataRow[] selectedRows = dataTableMergeData.Select(textBoxKeyField.Text + " ='" + key.Value.keyCol.ToString() + "'");

                        foreach (Word.MailMergeField mailMergeField in mailMerge.Fields)
                        {
                            if (mailMergeField.Code.Text.IndexOf(" MERGEFIELD " + "FirstName" + " ") > -1)
                            {
                                mailMergeField.Select();
                                
                                mailMerge.Application.Selection.TypeText(selectedRows[1][2].ToString());
                            }
                            else if (mailMergeField.Code.Text.IndexOf(" MERGEFIELD " + "Product_name" + " ") > -1)
                            {
                                mailMergeField.Select();
                                mailMerge.Application.Selection.TypeText(selectedRows[1][4].ToString());
                                for (int i = 0; i < selectedRows.Length; i++)
                                {
                                    if (i < selectedRows.Length - 1)
                                       

                                    {
                                        for (int tableIndex = 0; tableIndex < newDocument.Tables.Count; tableIndex++)
                                        {

                                        }
/*                                        foreach  (Word.Table table in wordDocument.Tables)
                                        {
                                            foreach (Word.Row row in table.Rows)
                                            {
                                                foreach (Word.Cell cell in row.Cells)
                                                {
                                                    
                                                }
                                            }
                                            
                                        }*/
                                        mailMerge.Application.Selection.InsertAfter("\r\n" + selectedRows[i + 1][4].ToString());
                                    }

                                }

                            }
                        }
                        
                        mailItem.Subject = "Test";
                        mailItem.To = "mohammed.saif.ibrahim@in.abb.com";
                        mailItem.Body = newDocument.Content.Text;
                        mailItem.Send();
                        newDocument.Close(false);
                        wordApplication.Quit(false);
                        Marshal.ReleaseComObject(newDocument);
                        Marshal.ReleaseComObject(wordApplication);
                    }
                }
            }

        }

        private void buttonAddRemoveKeyField_Click(object sender, EventArgs e)
        {
            if (buttonAddRemoveKeyField.Text == "Add >>" && listBoxFields.SelectedItem != null)
            {
                textBoxKeyField.Text = listBoxFields.SelectedItem.ToString();
                listBoxFields.Items.RemoveAt(listBoxFields.SelectedIndex);
                buttonAddRemoveKeyField.Text = "<< Remove";
                WriteDocumentProperty("KeyField", textBoxKeyField.Text);
            }
            else
            {
                listBoxFields.Items.Add(textBoxKeyField.Text);
                textBoxKeyField.Text = "";
                buttonAddRemoveKeyField.Text = "Add >>";
                DeleteDocumentProperty("KeyField");

            }
        }

        private void tabControlAdvancedMerge_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlAdvancedMerge.SelectedIndex == 1)
            {
                btnNext_Click(sender, e);
            }
            else if (tabControlAdvancedMerge.SelectedIndex == 0)
            {
                btnNext.Text = "Next";

            }
        }

        private void buttonAddChildField_Click(object sender, EventArgs e)
        {
            if (listBoxFields.SelectedItem != null)
            {
                listBoxChildFields.Items.Add(listBoxFields.SelectedItem);
                WriteDocumentProperty("ChildFields", listBoxFields.SelectedItem.ToString());
                listBoxFields.Items.RemoveAt(listBoxFields.SelectedIndex);
            }
            
        }

        private void buttonRemoveChildField_Click(object sender, EventArgs e)
        {
            if (listBoxChildFields.SelectedItem != null)
            {
                listBoxFields.Items.Add(listBoxChildFields.SelectedItem);
                DeleteDocumentProperty("ChildFields", listBoxChildFields.SelectedItem.ToString());
                listBoxChildFields.Items.RemoveAt(listBoxChildFields.SelectedIndex);

            }
        }
    }
}
