using System;
using System.IO;
using System.Windows.Forms;
using System.Security;
using System.Data;
using System.Data.OleDb;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace ManyToOne
{

    public partial class FormDataSource : Form
    {
        string filePath = string.Empty;
        string fileExt = string.Empty;


        public FormDataSource()
        {
            InitializeComponent();
        }

        private void LoadData()
        {
            string dataSheet = string.Empty;
            string dataSheets = string.Empty;
            char[] separotor = { '|' };
            string[] dataSheetsList;

            if (ReadDocumentProperty("DataSourcePath") != null)
            {
                filePath = ReadDocumentProperty("DataSourcePath");
            }
            if (ReadDocumentProperty("DataSourceSheets") != null)
            {
                dataSheets = ReadDocumentProperty("DataSourceSheets");
                dataSheetsList = dataSheets.Split(separotor, StringSplitOptions.RemoveEmptyEntries);
                if (dataSheetsList != null)
                {
                    comboBoxSheets.Items.Add("Select A Worksheet");
                    foreach (string item in dataSheetsList)
                    {
                        comboBoxSheets.Items.Add(item);
                    }

                }
            }
            if (ReadDocumentProperty("DataSourceSheet") != null)
            {
                dataSheet = ReadDocumentProperty("DataSourceSheet");
            }
            labelPath.Text = filePath;
            if (dataSheet != string.Empty)
            {
                //comboBoxSheets.Items.Add(dataSheet);
                comboBoxSheets.SelectedItem = dataSheet;
            }

            if (comboBoxSheets.SelectedItem != null)
            {

                //Get the worksheet data
                DataTable dataTable = OpenFile();
                if (dataTable != null)
                {
                    dataGridViewDataSource.DataSource = dataTable;

                }

            }
            /*else
            {
                MessageBox.Show("Please select a worksheet", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBoxSheets.Focus();
            }*/
        }


        void WriteDocumentProperty(string propertyName, string propertyValue)
        {
            Office.DocumentProperties properties;
            properties = Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

            if (propertyName == "DataSourceSheets" && ReadDocumentProperty("DataSourceSheets")!= null)
            {
                propertyValue = ReadDocumentProperty(propertyName) + "|" + propertyValue;
            }

            if (ReadDocumentProperty(propertyName) != null)
            {
              properties[propertyName].Delete();
            }
            
            properties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue);
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
                /*foreach (Office.DocumentProperty prop in properties)
                {
                    if (prop.Name == propertyName)
                    {
                        return prop.Value.ToString();
                    }
                }*/
            }

            return null;
        }

        public DataTable GetSheets()
        {
            string conn = string.Empty;
            DataTable dtSheets = null;

            fileExt = string.Empty;

            if (openFileDialogDataSource.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialogDataSource.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    if (fileExt.CompareTo(".xls") == 0)
                        conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007
                    else
                        conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007 

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
                }
                labelPath.Text = filePath;
            }

            return dtSheets;

        }


        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtSheets = null;
            DataTable dtexcel = new DataTable();
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

                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from ["+ comboBoxSheets.SelectedItem.ToString() + "]", con); //here we read data from sheet1  
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
            DataTable dataTable = new DataTable();
            try
                {
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                    //Read worksheet into data table
                        dataTable = ReadExcel(filePath, fileExt);
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
        private void btnDataSource_Click(object sender, EventArgs e)
        {
            DataTable workSheets = GetSheets();
            if (workSheets != null)
            {
                comboBoxSheets.Items.Clear();
                comboBoxSheets.Items.Add("Select A Worksheet");
                foreach (DataRow sheet in workSheets.Rows)
                {
                    comboBoxSheets.Items.Add(sheet["TABLE_NAME"].ToString());
                    WriteDocumentProperty("DataSourceSheets", sheet["TABLE_NAME"].ToString());
                }
                comboBoxSheets.SelectedIndex = 0;
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonSetDataSource_Click(object sender, EventArgs e)
        {
            Object missing = System.Reflection.Missing.Value;
            if (comboBoxSheets.SelectedIndex != 0 && filePath != null)
            {
                //save selected data source to document properties
                WriteDocumentProperty("DataSourcePath", filePath);
                WriteDocumentProperty("DataSourceSheet", comboBoxSheets.SelectedItem.ToString());


                //Get the worksheet data
                DataTable dataTable = OpenFile();
                if (dataTable != null)
                {
                    dataGridViewDataSource.DataSource = dataTable;

                }
                
                //update word document's data source to the selected data source
                try
                {
                    Globals.ThisAddIn.Application.ActiveDocument.MailMerge.OpenDataSource(filePath, missing, missing, missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, "select * from [" + comboBoxSheets.SelectedItem.ToString() + "]", missing, missing, missing);

                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                

            }
            else
            {
                MessageBox.Show("Please select a worksheet","Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBoxSheets.Focus();
            }
            
        }

        private void FormDataSource_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void comboBoxSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxSheets.SelectedIndex != 0)
            {
                buttonSetDataSource.Enabled = true;
            }
            else
            {
                buttonSetDataSource.Enabled = false;
            }
        }
    }
}
