using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.IO;
using System.Collections;
using ExcelDataReader;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace QSKSKS
{
    public partial class FormLoadExcel : Form
    {
        Microsoft.Office.Interop.Excel.Application oXA = new Microsoft.Office.Interop.Excel.Application();
        

        public string pathBOQ = "C:\\TTD BIM 5D\\txt\\BOQPath.txt";
        public FormLoadExcel()
        {
            InitializeComponent();
        }
        public FormLoadExcel(string _cateName)
        {
            InitializeComponent();
            this.cateName = _cateName;
            //this.linkPath = _linkPath;
        }
        public void closeApplication(Microsoft.Office.Interop.Excel.Application oXA)
        {
            if (oXA != null && !oXA.Visible)
            {
                if (oXA != null)
                {
                    oXA.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXA);
                    oXA = null;
                }
                System.Diagnostics.
                Process[] Processes;
                Processes = System.Diagnostics.
                Process.GetProcessesByName("EXCEL.EXE");
                foreach (System.Diagnostics.Process p in Processes)
                {
                    if (p.MainWindowTitle.Trim() == "")
                        p.Kill();
                }
            }
        }
        public string linkPath { get; set; }
        public string cateName { get; set; }
        public string getRow
        {
            get{return tbSelectRow.Text;}
        }
        private void FormLoadExcel_Load(object sender, EventArgs e)
        {
            tbCategory.Text = cateName;
            //tbLinkPath.Text = linkPath;
        }
        
        private void bOK_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }
        private void bGet_Click(object sender, EventArgs e)
        {
            //Mở file Excel - lấy data vào dgv - kích đúp chuột vào dòng cần tìm - ghi vào tbSelectRow.Text
            //nút OK là để ghi tbSelectRow.Text (ở FormLoadExcel) vào tbGetRow.Text(ở FormBOQ) và close() FormLoadExcel lại
            selectRow(pathBOQ, dgvExcel);
        }
        private void cmtInsert_Opening(object sender, CancelEventArgs e)
        {
            int rowIndex = dgvExcel.CurrentRow.Index;
            var indexValue = dgvExcel.Rows[rowIndex].Cells[0].Value;
            tbSelectRow.Text = Convert.ToString(indexValue);
        }

        private string _filePath;
        private int _findRow;
        public string filePath
        {
            get => _filePath;
            set { _filePath = value; }
        }
        public int findRow
        {
            get => _findRow;
            set { _findRow = value; }
        }
        //DataTableCollection tableCollection;
        public void selectRow(string filePath, DataGridView dgv)
        {
            //paths is link (file .txt) that contains the path to Excel file
            StreamReader txt = new StreamReader(filePath);
            string myFile = txt.ReadToEnd();
            txt.Close();
            var rootPath = Path.GetFullPath(myFile);
            if (System.IO.File.Exists(myFile))
            {
                //Open Excel File to View
                Object misValue = System.Reflection.Missing.Value;
                Workbook wbBOQ = oXA.Application.Workbooks.Open(rootPath, misValue, false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Worksheet shBOQ = wbBOQ.Worksheets["BOQ"];
                Excel.Range rng = shBOQ.UsedRange;
                int lrBOQ = rng.Rows.Count;
                int lcBOQ = rng.Columns.Count;
                int soCotCanLay = 4; //Thay soCotCanLay bằng bao nhiêu để lấy số cột cần lấy ra
                dgv.ColumnCount = soCotCanLay;
                for (int r = 1; r < lrBOQ; r++)
                {
                    if (shBOQ.Cells[r,1].Value2 != null)
                        {
                        String[] rowData = new String[lcBOQ];
                        rowData[0] = Convert.ToString(r);
                        for (int i = 1; i < soCotCanLay; i++) 
                        {
                            rowData[i] = Convert.ToString(rng.Cells[r, i + 1].Value2);
                        }
                        dgv.Rows.Add(rowData);
                    }
                }
                wbBOQ.Close();
                oXA.Quit();
                closeApplication(oXA);
                #region Draft raw code
                //-----------------------------------------------------------------------------------------------------------------------------------------------
                //using (var streamBOQ = File.Open(myFile, FileMode.Open, FileAccess.Read))
                //{
                //    using(IExcelDataReader reader=ExcelReaderFactory.CreateReader(streamBOQ))
                //    {
                //        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                //        {
                //            ConfigureDataTable=(_)=>new ExcelDataTableConfiguration() {  UseHeaderRow = true }
                //        });
                //        tableCollection = result.Tables;
                //        cboSheet.Items.Clear();
                //        foreach (Excel.DataTable table in tableCollection)
                //        {
                //            dgv.DataSource = table;
                //        }
                //    }
                //    streamBOQ.Close();
                //}
                //-----------------------------------------------------------------------------------------------------------------------------------------------
                //                string excelFile = "SELECT * FROM Authors";
                //                OleDbConnection theConnection = new OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;
                //data source='"+myFile+"';Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"");
                //                OleDbDataAdapter dtAdapter = new OleDbDataAdapter(excelFile, theConnection);
                //                DataSet ds = new DataSet();
                //                theConnection.Open();
                //                dtAdapter.Fill(ds, "BOQ");
                //                theConnection.Close();
                //                dgv.DataSource = ds;
                //                dgv.DataMember = "BOQ";
                //-----------------------------------------------------------------------------------------------------------------------------------------------
                #endregion
            }
            else
            {
                MessageBox.Show("Error on FormLoadExcel, when you assign a row to fill data!" + "\n" + "Unable to open Workbook because of BOQ file uncorrect path." + "\n" + " Please check Import Form again!!",
                "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void dgvExcel_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = dgvExcel.CurrentRow.Index;
            var indexValue = dgvExcel.Rows[rowIndex].Cells[0].Value;
            tbSelectRow.Text = Convert.ToString(indexValue);
        }
        private void dgvExcel_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = dgvExcel.CurrentRow.Index;
            var indexValue = dgvExcel.Rows[rowIndex].Cells[0].Value;
            tbSelectRow.Text = Convert.ToString(indexValue);
        }
    } 
}
