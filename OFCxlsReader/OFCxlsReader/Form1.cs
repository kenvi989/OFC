using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using InDesign;

namespace OFCxlsReader
{
    public partial class Form1 : Form
    {
        public InDesign.Application _application = null;
        public Document _docmument;
        public Form1()
        {
            InitializeComponent();
            LoadTemplate();
        }

        DataTableCollection tableCollection;

        private void button1_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog openFileDialog=new OpenFileDialog() { Filter="Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    using(var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable=(_)=>new ExcelDataTableConfiguration() { UseHeaderRow = true}
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);//add sheet to combobox
                        }
                    }
                }
            }
        }

        private void CreateTable(DataSet dataSet)
        {
            DataTable newTable = new DataTable("table");
            newTable.Columns.Add("出発都市3コード", typeof(string));
            newTable.Columns.Add("出発都市3コード");
            dataSet.Tables.Add(newTable);
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private bool DoScripts(string jsxFile)
        {
            try
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(jsxFile);
                _application.DoScript(jsxFile, idScriptLanguage.idJavascript);
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("ErrorExecuteJSX　エラー:　" + ex.Message, "通知", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                MessageBox.Show("テンプレート、または紙面体裁テキストファイル名等のご確認をお願い致します。" + ex.Message);
                return false;
            }
        }

        private void LoadTemplate()
        {
            try
            {
                List<string> TemplateList = new List<string>();
                string folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath);
                if (Directory.Exists(folder + "\\Template"))
                    folder = folder + "\\Template";
                else if (Directory.Exists(folder + "\\template"))
                    folder = folder + "\\template";
                string[] tempFiles = Directory.GetFiles(folder, "*.indd");
                foreach (string templates in Directory.GetFiles(folder, "*.indd"))
                { TemplateList.Add(templates); }
                foreach (string items in TemplateList)
                {
                    cmbTemplate.Items.Add(System.IO.Path.GetFileName(items));
                }
            }
            catch (Exception exp) { }
        }

        private Document OpenDocument(string filePath)
        {
            Document documentTemplate = null;
            string fileName = System.IO.Path.GetFileName(filePath);

            //Open document
            if (documentTemplate == null)
            {
                documentTemplate = _application.Open(filePath, false);
            }
            // Global variable setting

            this._docmument = documentTemplate;
            return documentTemplate;
        }
        private InDesign.Application OpenApplication()
        {
            try
            {
                Type type2018 = Type.GetTypeFromProgID("InDesign.Application.CC.2018");
                Type type2019 = Type.GetTypeFromProgID("InDesign.Application.CC.2019");
                Type type = Type.GetTypeFromProgID("InDesign.Application");
                if (type2018 == null & type2019 == null & type == null)
                {
                    throw new Exception("Adobe InDesign CC is not installed");
                    return null;
                }
                else if (type2018 != null) _application = (InDesign.Application)Activator.CreateInstance(type2018);
                else if (type2019 != null) _application = (InDesign.Application)Activator.CreateInstance(type2019);
                else _application = (InDesign.Application)Activator.CreateInstance(type);
                //app = COMCreateObject("InDesign.Application");
                if (_application == null)
                {
                    if (type2018 != null) _application = (InDesign.Application)Activator.CreateInstance(type2018);
                    else if (type2019 != null) _application = (InDesign.Application)Activator.CreateInstance(type2019);
                    else _application = (InDesign.Application)Activator.CreateInstance(type);
                    System.Threading.Thread.Sleep(10000);
                }
                // app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;               
                return _application;
            }
            catch (Exception ex)
            {
                //log.Error(ex);
                return null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Create a StringBuilder that expects to hold 50 characters.
            StringBuilder sb = new StringBuilder();

            //string bfShuppatsu = "";
            string bfChaku = "";
            string bfKaisha = "";


            List<string[]> lstJuryo = new List<string[]>();
            bool findJuryo = false;
            List<string> lstKaisha = new List<string>();
            bool findKaisha = false;

            DataTable table = tableCollection[cboSheet.SelectedItem.ToString()];
            foreach (DataRow row in table.Rows)
            {

                if (bfChaku == row["到着都市3コード"].ToString())
                {
                    //adilhan hotiin medeelliin davhtsaliig shalgaj bgaa heseg
                    if (bfKaisha != row["航空会社コード"].ToString())
                    {
                        lstKaisha.Add(row["航空会社コード"].ToString());
                    }

                    findJuryo = false;
                    foreach (string[] ju in lstJuryo)
                    {
                        if (ju[0] == row["重量"].ToString())
                        {
                            findJuryo = true;
                        }
                    }

                    if (findJuryo != true)
                    {
                        string[] ryoandprice = { row["重量"].ToString() , row["賃率"].ToString() };
                        lstJuryo.Add(ryoandprice);
                    }
                }
                else
                {
                    //group bolgoson hotiin medeeleliig hevleh heseg
                    string strkaisha = " ";
                    foreach (string kai in lstKaisha)
                    {
                        if(strkaisha == " ")
                        {
                            strkaisha = strkaisha + kai;
                        }
                        else
                        {
                            strkaisha = strkaisha + "," + kai;
                        }
                    }

                    findKaisha = false;
                    foreach (string[] ju in lstJuryo)
                    {
                        if (findKaisha == false)
                        {
                            sb.AppendLine(strkaisha + "                 " + ju[0] + "      " + ju[1] + " | ");
                            findKaisha = true;
                        }
                        else
                        {
                            sb.AppendLine("                 " + ju[0] + "      " + ju[1] + " | ");
                        }
                    }

                    //shine hotiin medeelel avch ehelj bgaa heseg
                    lstJuryo = new List<string[]>();
                    lstKaisha = new List<string>();
                    string[] ryoandprice = {row["重量"].ToString(), row["賃率"].ToString()};
                    lstJuryo.Add(ryoandprice);
                    lstKaisha.Add(row["航空会社コード"].ToString());

                    //shine hot bolon ulsiin ner hevelj bgaa heseg
                    sb.AppendLine(row["到着都市"].ToString() + " (" + row["到着都市3コード"].ToString() + ")      " + row["到着国"].ToString() + " | ");

                }

                bfChaku = row["到着都市3コード"].ToString();
                bfKaisha = row["航空会社コード"].ToString();
            }
        

            using (StreamWriter file = new StreamWriter(@"D:\Bold\Projects\myofc\hereIam.txt"))
            {
                file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
            }
        }
    }
}