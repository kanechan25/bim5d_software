using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class FormBOQ : Form
    {

        public string pathACate = "C:\\TTD BIM 5D\\txt\\ListACate.txt";
        public string pathALevel = "C:\\TTD BIM 5D\\txt\\ListALevel.txt";
        public string pathARowLevel = "C:\\TTD BIM 5D\\txt\\ListARowLevel.txt";
        public string pathAEle = "C:\\TTD BIM 5D\\txt\\ListAElement.txt";
        public string pathARowEle = "C:\\TTD BIM 5D\\txt\\ListARowEle.txt";
        public string pathAEleDep = "C:\\TTD BIM 5D\\txt\\ListAEleDep.txt";
        public string pathSCate = "C:\\TTD BIM 5D\\txt\\ListSCate.txt";
        public string pathSLevel = "C:\\TTD BIM 5D\\txt\\ListSLevel.txt";
        public string pathSRowLevel = "C:\\TTD BIM 5D\\txt\\ListSRowLevel.txt";
        public string pathSEle = "C:\\TTD BIM 5D\\txt\\ListSElement.txt";
        public string pathSRowEle = "C:\\TTD BIM 5D\\txt\\ListSRowEle.txt";
        public string pathSEleDep = "C:\\TTD BIM 5D\\txt\\ListSEleDep.txt";
        public string  saveListACate = "C:\\TTD BIM 5D\\txt\\saveListACate.txt";    //Hiện nay chưa dùng đến
        public string  saveListSCate = "C:\\TTD BIM 5D\\txt\\saveListSCate.txt";    //Hiện nay chưa dùng đến
        public string  pathBOQ = "C:\\TTD BIM 5D\\txt\\BOQPath.txt";

        FormSupport frmSupport = new FormSupport();
        WaitFormFunction waitForm = new WaitFormFunction();
        public FormBOQ()
        {
            InitializeComponent();
        }
        private void FormBOQ_Load(object sender, EventArgs e)
        {
            //1. UI
            for (int i = 0; i < 7; i++)
            {dgvTask.Rows.Add();}
           
            bFillL.Enabled = false;
            bFillN.Enabled = false;
            cbFill.Enabled = false;
            ckbFill.Checked = false;
            bGetQty.Enabled = true;
            cbGet.Enabled = true;
            ckbGet.Checked = true;
            bSelectRow.Enabled = false;
            //2. Load data to combobox
            GetDataToCBX(pathSCate,cbCate);
            GetDataToCBXLevel(pathSLevel, cbLevel);
        }

        #region Checked Change to Click Structure or Architecture
        private void rbCheckStr_CheckedChanged(object sender, EventArgs e)
        {
            cbCate.Items.Clear();
            cbLevel.Items.Clear();
            //dgvListCate2.Rows.Clear();
            if (rbCheckArch.Checked == true)
            {
                GetDataToCBX(pathACate, cbCate);
                GetDataToCBXLevel(pathALevel, cbLevel);
                //GetDataToDGV(pathACate, dgvListCate2);
            }
            else if (rbCheckStr.Checked == true)
            {
                GetDataToCBX(pathSCate, cbCate);
                GetDataToCBXLevel(pathSLevel, cbLevel);
                //GetDataToDGV(pathSCate, dgvListCate2);
            }
            cbLevel.Text = "All";
        }

        private void rbCheckArch_CheckedChanged(object sender, EventArgs e)
        {
            cbCate.Items.Clear();
            cbLevel.Items.Clear();
            //dgvListCate2.Rows.Clear();
            if (rbCheckArch.Checked == true)
            {
                GetDataToCBX(pathACate, cbCate);
                GetDataToCBXLevel(pathALevel, cbLevel);
                //GetDataToDGV(pathACate, dgvListCate2);
            }
            else if (rbCheckStr.Checked == true)
            {
                GetDataToCBX(pathSCate, cbCate);
                GetDataToCBXLevel(pathSLevel, cbLevel);
                //GetDataToDGV(pathSCate, dgvListCate2);
            }
            cbLevel.Text = "All";
        }
        #endregion

        #region Checked Change Fill and Get Quantity + SpeialColNum
        private void ckbFill_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbFill.Checked == true)
            {
                bFillL.Enabled = true;
                bFillN.Enabled = true;
                cbFill.Enabled = true;
                bGetQty.Enabled = false;
                cbGet.Enabled = false;
                ckbGet.Checked = false;
                if (cbFill.Text == "Fill with a Selective Row")
                {bSelectRow.Enabled = true;}
            }
            else
            {
                bFillL.Enabled = false;
                bFillN.Enabled = false;
                cbFill.Enabled = false;
                bGetQty.Enabled = true;
                cbGet.Enabled = true;
                ckbGet.Checked = true;
            }
        }
        private void ckbGet_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbGet.Checked == true)
            {
                bFillL.Enabled = false;
                bFillN.Enabled = false;
                cbFill.Enabled = false;
                bGetQty.Enabled = true;
                cbGet.Enabled = true;
                ckbFill.Checked = false;
                bSelectRow.Enabled = false;
            }
            else
            {
                bFillL.Enabled = true;
                bFillN.Enabled = true;
                cbFill.Enabled = true;
                bGetQty.Enabled = false;
                cbGet.Enabled = false;
                ckbFill.Checked = true;
            }
        }
        public int SpecialColNum(Worksheet wsh, string str, int orderNum)
        //Find the position (order Number) column letter that contain a special character, appeared for the n time (calculate on cell by cell value)
        {
            Excel.Range rng = wsh.UsedRange;
            int lcBOQ = rng.Columns.Count;
            int count = 0;
            int col = 100;
            for (int i = 1; i <= lcBOQ; i++)
            {
                string cellValue = wsh.Cells[1, i].Value2;
                if (cellValue != null)
                {
                    if (cellValue.Contains(str) == true)
                    {
                        count++;
                        if (count == orderNum)
                        { col = i; }
                    }
                }
            }
            return col;
        }
        #endregion

        #region Button OK_Click + Form_Close+ closeApplication+Check BOQ Button + NumColRange2

        private void bOK_Click(object sender, EventArgs e)
        {
            waitForm.Close();
            frmSupport.Close();
            this.Close();
        }
        private void FormBOQ_Closed(object sender, FormClosedEventArgs e)
        {
            
        }
        public void closeApplication(Microsoft.Office.Interop.Excel.Application oXL)
        {
            if (oXL != null && !oXL.Visible)
            {
                if (oXL != null)
                {
                    oXL.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                    oXL = null;
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

        private void bCheck_Click(object sender, EventArgs e)
        { CheckBOQFile(); }
        public void CheckBOQFile()
        //Check BOQ sheet Exist & Open BOQ file
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                //1. Open file BOQ Excel
                string myPathBOQ = @"C:\TTD BIM 5D\txt\BOQPath.txt";
                StreamReader txt = new StreamReader(myPathBOQ);
                string myFile = txt.ReadToEnd();
                txt.Close();
                var rootPath2 = Path.GetFullPath(myFile);
                if (System.IO.File.Exists(myFile))
                {
                    //1. Visible Excel File to View
                    Workbook oWB = oXL.Workbooks.Open(rootPath2);
                    //DoesSheetExists("Format", oWB);
                    //frmSupport.DoesSheetExists("BOQ", oWB); //Check xem có sheet BOQ không, không có thì tạo | Function "DoesSheetExists" will clear all cell in worksheet
                    oXL.Visible = true;
                }
                else
                {
                    MessageBox.Show(" Error on FormBOQ, when you use the Check BOQ file button." + "\n" + " Unable to open Workbook because of BOQ file uncorrect path." + "\n" + " Please check Import Form again!!",
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    closeApplication(oXL);
                }
            }
            catch (Exception)
            {
                closeApplication(oXL);
                MessageBox.Show(" Error on FormBOQ, when you use the Check BOQ file button." + "\n" + " Unable to open Workbook because of uncorrect path." + "\n" + " Please check Import Form again!!",
                        "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public int NumColRange2(Worksheet wsheet, string colName)
        {
            int colNum = 1;
            for (int iCol = 1; iCol < 25; iCol++)
            {
                string cellValue = wsheet.Cells[1, iCol].Value2;
                if (cellValue == colName)
                {
                    colNum = iCol;
                    return colNum;
                }
            }
            return colNum;
        }        public int NumColRangeBOQ(Worksheet wsheet, string colName, int numCol)
        {
            int colNum = 1;
            for (int i = 1; i <= numCol; i++)
            {
                for (int iCol = 1; iCol < 25; iCol++)
                {
                    string cellValue = wsheet.Cells[i, iCol].Value2;
                    if (cellValue == colName)
                    {
                        colNum = iCol;
                        return colNum;
                    }
                }
            }

            return colNum;
        }

        #endregion

        #region Save DataGridView, GetData to ComboBox, to dataGridview, SelectRow_Click
        public void saveDGV(string myPathTXT, DataGridView dgv)
        {
            TextWriter txt = new StreamWriter(myPathTXT);
            for (int rc = 0; rc < (dgv.Rows.Count) - 1; rc++)
            {
                txt.WriteLine(dgv.Rows[rc].Cells[0].Value.ToString());
            }
            txt.Close();
        }
        public void GetDataToCBX(string pathFile, ComboBox cbx)
        {
            cbx.Items.Clear();
            string line = "";
            StreamReader filePath = new StreamReader(pathFile);
            while (line != null)
            {
                line = filePath.ReadLine();
                if (line != null)
                {
                    int ld = 0;
                    cbx.Items.Add(line);
                    ld++;
                }
            }
            filePath.Close();
        }
        public void GetDataToCBXLevel(string pathFile, ComboBox cbx)
        {
            cbx.Items.Clear();
            cbLevel.Items.Add("All");
            string line = "";
            StreamReader filePath = new StreamReader(pathFile);
            while (line != null)
            {
                line = filePath.ReadLine();
                if (line != null)
                {
                    int ld = 0;
                    cbx.Items.Add(line);
                    ld++;
                }

            }
            filePath.Close();
        }
        public void GetDataToDGV(string pathFile, DataGridView dgvGet)
        {
            dgvGet.Rows.Clear();
            string line = "";
            StreamReader filePath = new StreamReader(pathFile);
            while (line != null)
            {
                line = filePath.ReadLine();
                if (line != null)
                {
                    int ld = 0;
                    dgvGet.Rows.Add(line);
                    ld++;
                }
            }
            filePath.Close();
        }
        private void bSelectRow_Click(object sender, EventArgs e)
        {
            try
            {
                using (FormLoadExcel frm = new FormLoadExcel(cbCate.Text))
                //Muốn truyền tham số nào sang bên form kia thì thêm vào trên đây (ví dụ cbCate.Text, myPathS)
                {
                    if (frm.ShowDialog() == DialogResult.OK)
                        tbGetRow.Text = frm.getRow;
                    frm.cateName = cbCate.Text;
                    //frm.linkPath = myPathS;
                }
            }
            catch (Exception)
            {
                MessageBox.Show(" Error on FormBOQ, when you use the Select Row button for the Structural file." + "\n" + "Unable to open Workbook because of uncorrect path." + "\n" + " Please check Import Form again!!",
                        "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region Haven't programmed yet!!
        private void bGetQty_Click(object sender, EventArgs e)
        {

            string CateValue = cbCate.Text;
            if (cbCate.Text == "" | cbGet.Text == "")
            {
                MessageBox.Show("You have to chose a value from" + "\n" + " \"Categories\" and \"Get Quantity\" ComboBox, please! ", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    // Case 1: Structure File
                    if (rbCheckStr.Checked == true)
                    {
                        Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                        string myPathS = @"C:\TTD BIM 5D\txt\SCatePath.txt";
                        if (frmSupport.checkLinkPath(myPathS) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                        {
                            //waitForm.Show(this);
                            //Declare 2 Worksheet : BOQ and Category (contain categories we chose)
                                //1.1.File Category, copy Source, Sheet Cần chọn
                                    StreamReader txt2 = new StreamReader(myPathS);
                                    string linkFileCopy = txt2.ReadToEnd();
                                    txt2.Close();
                                    var rootCopyPath = Path.GetFullPath(linkFileCopy);
                                    Workbook wbCateCopy = oXL2.Workbooks.Open(rootCopyPath);
                                    Worksheet wsCateCopy = wbCateCopy.Worksheets[CateValue];
                                    Excel.Range sourceRng = wsCateCopy.UsedRange;
                                    sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();
                            //1.2.File BOQ, Sheet InputData
                                    StreamReader txt = new StreamReader(pathBOQ);
                                    string linkFilePaste = txt.ReadToEnd();
                                    txt.Close();
                                    var rootPastePath = Path.GetFullPath(linkFilePaste);
                                    Workbook wbBOQPaste = oXL.Workbooks.Open(rootPastePath);
                                    frmSupport.DoesSheetExists("InputData", wbBOQPaste);
                                    Worksheet wsInputData = wbBOQPaste.Worksheets["InputData"];
                                    wsInputData.Cells.ClearContents();
                                    Worksheet wsBOQ = wbBOQPaste.Worksheets["BOQ"];
                                    Excel.Range BOQRng = wsInputData.get_Range("A1");
                                    BOQRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                            //Close file Category
                            Clipboard.Clear();
                            wbCateCopy.Close();
                            oXL2.Quit();
                            closeApplication(oXL2);
                            switch (cbGet.Text)
                            {
                                case "Get by Name & Level":
                                    GetbyName_Level(wsInputData);
                                    break;
                                case "Get by Level":
                                    // code block
                                    break;
                                case "Get by Name":
                                    // code block
                                    break;
                            }
                            oXL.DisplayAlerts = false;
                            wbBOQPaste.Save();
                            oXL.DisplayAlerts = true;
                            //oXL.Visible = true;
                            oXL.Quit();
                            closeApplication(oXL);
                            //waitForm.Close();
                        }
                        else
                        {MessageBox.Show("You have not imported or not correct link Structural Category file path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);}
                    }
                    // Case 2: Architecture File
                    else if (rbCheckArch.Checked == true)
                    {
                        Microsoft.Office.Interop.Excel.Application oXL_ = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Application oXL2_ = new Microsoft.Office.Interop.Excel.Application();
                        string myPathA = @"C:\TTD BIM 5D\txt\ACatePath.txt";
                        if (frmSupport.checkLinkPath(myPathA) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                        {
                            //waitForm.Show(this);
                            //Declare 2 Worksheet : BOQ and Category (contain categories we chose)
                            //1.1.File Category, copy Source, Sheet Cần chọn
                            StreamReader txt2 = new StreamReader(myPathA);
                            string linkFileCopy = txt2.ReadToEnd();
                            txt2.Close();
                            var rootCopyPath = Path.GetFullPath(linkFileCopy);
                            Workbook wbCateCopy = oXL2_.Workbooks.Open(rootCopyPath);
                            Worksheet wsCateCopy = wbCateCopy.Worksheets[CateValue];
                            Excel.Range sourceRng = wsCateCopy.UsedRange;
                            sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();
                            //1.2.File BOQ, Sheet InputData
                            StreamReader txt = new StreamReader(pathBOQ);
                            string linkFilePaste = txt.ReadToEnd();
                            txt.Close();
                            var rootPastePath = Path.GetFullPath(linkFilePaste);
                            Workbook wbBOQPaste = oXL_.Workbooks.Open(rootPastePath);
                            frmSupport.DoesSheetExists("InputData", wbBOQPaste);
                            Worksheet wsInputData = wbBOQPaste.Worksheets["InputData"];
                            wsInputData.Cells.ClearContents();
                            Worksheet wsBOQ = wbBOQPaste.Worksheets["BOQ"];
                            Excel.Range BOQRng = wsInputData.get_Range("A1");
                            BOQRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                            //Close file Category
                            Clipboard.Clear();
                            wbCateCopy.Close();
                            oXL2_.Quit();
                            closeApplication(oXL2_);
                            switch (cbGet.Text)
                            {
                                case "Get by Name & Level":
                                    GetbyName_Level(wsInputData);
                                    break;
                                case "Get by Level":
                                    // code block
                                    break;
                                case "Get by Name":
                                    // code block
                                    break;
                            }
                            wbBOQPaste.Save();
                            //oXL_.Visible = true;
                            oXL_.Quit();
                            closeApplication(oXL_);
                            //waitForm.Close();
                        }
                        else
                        { MessageBox.Show("You have not imported or not correct link Architectural Category file path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void bCreate_Click(object sender, EventArgs e)
        {

        }

        private void bFillN_Click(object sender, EventArgs e)
        {
            //Thông báo nhắc Close file BOQ trc khi điền KL vào file BOQ đó

        }

        private void bFillL_Click(object sender, EventArgs e)
        {
            //Thông báo nhắc Close file BOQ trc khi điền KL vào file BOQ đó
            try
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure? " + "\n" + "You will start to fill Quantity for "
                    + cbCate.Text + " category. " + "\n" + "You can't stop and it will take serveral minutes?"
                    + "\n" + "Continue??", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dialogResult == DialogResult.No)
                { }
                else
                {
                    if (dgvBOQSA.Rows.Count == 0)
                    {
                        MessageBox.Show("You have nothing to insert to BOQ file." + "\n" + "Please Load Quantity and QTO then Fill Quantity. ",
                          "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    }
                    else
                    {
                        //Insert data to Excel
                        Microsoft.Office.Interop.Excel.Application oXL_ = new Microsoft.Office.Interop.Excel.Application();
                        StreamReader txt = new StreamReader(pathBOQ);
                        string linkFilePaste = txt.ReadToEnd();
                        txt.Close();
                        var rootBOQPath = Path.GetFullPath(linkFilePaste);
                        Workbook wbBOQ = oXL_.Workbooks.Open(rootBOQPath);
                        Worksheet wshBOQ = wbBOQ.Worksheets["BOQ"];
                        int colName = NumColRangeBOQ(wshBOQ, "Task Name", 3);
                        int colUnit = NumColRangeBOQ(wshBOQ, "Unit", 3);
                        int colQty = NumColRangeBOQ(wshBOQ, "Quantity", 3);
                        //MessageBox.Show("Cot Nam la " + colName + " , Cot Unit la " + colUnit + " , Cot Qty la " + colQty);
                        int insertRow = Convert.ToInt32(tbGetRow.Text);
                        int numData = dgvBOQSA.Rows.Count;
                        for (int row = 0; row < numData; row++)
                        {


                            wshBOQ.Rows[insertRow].Insert();
                            wshBOQ.Cells[insertRow + 1, 2].Value = dgvBOQSA.Rows[row].Cells[1].Value;

                        }


                        oXL_.Visible = true;
                        // Save và Close app
                        //wbBOQ.Save();
                        //wbBOQ.Close();
                        //System.Runtime.InteropServices.Marshal.ReleaseComObject(wbBOQ);
                        //oXL_.Quit();
                        //System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL_);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        #endregion
        Dictionary<int, string> row_Name = new Dictionary<int, string>();
        Dictionary<int, List<string>> row_Subtotal = new Dictionary<int, List<string>>();
        List<string> listSubtotal = new List<string>();
        List<int> numColList = new List<int>();
        public void GetbyName_Level(Worksheet wsh)
        {
            row_Name.Clear();
            row_Subtotal.Clear();
            dgvBOQSA.Rows.Clear();
            numColList.Clear();
            Excel.Range rng = wsh.UsedRange;
            int lrBOQ = rng.Rows.Count;
            int lcBOQ = rng.Columns.Count;
            //1. Define All Column Name by Cell Values
            List<int> colRebar = FindRebarColumn(wsh);
            int colName = NumColRange2(wsh, "Name");
            int colLevel = NumColRange2(wsh, "Level");
            //Column Number List that included in 7 column Quantity of dgvBOQSA
            numColList.Add(SpecialColNum(wsh, "(", 1));
            numColList.Add(SpecialColNum(wsh, "(", 2));
            numColList.Add(SpecialColNum(wsh, "(", 3));
            numColList.Add(SpecialColNum(wsh, "(", 4));
            numColList.Add(Convert.ToInt32(colRebar[0]));
            numColList.Add(Convert.ToInt32(colRebar[1]));
            numColList.Add(Convert.ToInt32(colRebar[2]));
            //1. Get 2 Dict : 1 dict row : Name, 1 dict row : Quantity Subtotal
            int dgvBOQRow = 0;
            for (int n = 1; n <= lrBOQ; n++)
            {
                string nameValue = wsh.Cells[n, colName].Value2;
                string levelValue = wsh.Cells[n, colLevel].Value2;
                if (nameValue != "Subtotal" && nameValue != null && nameValue != "Name")
                {
                    if (nameValue.Contains("Deduct") == false && nameValue.Contains("Add") == false)
                    {
                        dgvBOQSA.Rows.Add();
                        listSubtotal.Clear();
                        int s = FindSubtotalRow(wsh, n);
                        for (int ncl = 0; ncl < numColList.Count; ncl++)
                        {listSubtotal.Add(Convert.ToString(wsh.Cells[s, numColList[ncl]].Value2));}
                        row_Name.Add(n, nameValue);
                        row_Subtotal.Add(s, listSubtotal);
                        //2. Load onto dgvBOQ
                        dgvBOQSA.Rows[dgvBOQRow].Cells[1].Value = nameValue;
                        dgvBOQSA.Rows[dgvBOQRow].Cells[2].Value = levelValue;
                        dgvBOQSA.Rows[dgvBOQRow].Cells[0].Value = dgvBOQRow + 1;
                        for (int dgv = 0; dgv < numColList.Count; dgv++)
                            {dgvBOQSA.Rows[dgvBOQRow].Cells[dgv+3].Value = listSubtotal[dgv];}
                        for (int ht = 0; ht < numColList.Count; ht++)
                            {dgvBOQSA.Columns[ht+3].HeaderText = Convert.ToString(wsh.Cells[1, numColList[ht]].Value2);}
                        dgvBOQRow++;
                    }
                }
            }
            //3. Load parameters onto dgvTask
            dgvTask.Rows.Clear();
            for (int i = 0; i < 7; i++)
            { dgvTask.Rows.Add(); }
            for (int ht = 0; ht < numColList.Count; ht++)
            {
                string parameterValue = wsh.Cells[1, numColList[ht]].Value2;
                if (parameterValue != null)
                {
                    dgvTask.Rows[ht].Cells[0].Value = parameterValue;
                    dgvTask.Rows[ht].Cells[1].Value = parameterValue;
                }
                else
                {}
            }
            //4. Calculate and Load Total Quantity onto dgvTask
            int rc = dgvBOQSA.Rows.Count;
            for (int q = 0; q < 7; q++)
            {
                if (dgvBOQSA.Columns[q+3].HeaderText != "")
                {
                    double sum = 0;
                    for (int r = 0; r < rc; r++)
                    {
                        double value = Convert.ToDouble(dgvBOQSA.Rows[r].Cells[q + 3].Value);
                        sum = sum + value;
                    }
                    dgvTask.Rows[q].Cells[2].Value = Convert.ToString(sum);
                }
            }
            //5. Remove duplicate rows of datagridview
        }
        private void bQTO_Click(object sender, EventArgs e)
        {
            removeDuplicate_Name_Level(dgvBOQSA, 1, 2);
            removeDuplicate_Name_Level(dgvBOQSA, 1, 2);
            fillOrderNumber(dgvBOQSA);

        }
        public void removeDuplicate_Name_Level(DataGridView dgv, int col1, int col2)
        {   //col1, col2 are column Index of datagridView
            try
            {
                for (int currentRow = 0; currentRow  < (dgv.Rows.Count - 1); currentRow ++)
                {
                    //1. Get Quantity from all cells of compare row
                    for (int compareRow = currentRow + 1; compareRow < dgv.Rows.Count; compareRow++)
                    {
                        if (((dgv.Rows[currentRow].Cells[col1].Value).Equals(dgv.Rows[compareRow].Cells[col1].Value)) &&
                            ((dgv.Rows[currentRow].Cells[col2].Value).Equals(dgv.Rows[compareRow].Cells[col2].Value)))
                        {
                            //2. If compare rows are equal to current rows
                            // Then cumulative summary of quantity to current rows, finally remove that duplicate rows (=compare rows)
                            for (int i = 0; i < 7; i++)
                            {
                                double lcurq = Convert.ToDouble(dgv.Rows[currentRow].Cells[i+3].Value);
                                double lcomq = Convert.ToDouble(dgv.Rows[compareRow].Cells[i + 3].Value);
                                double lsumq = lcurq + lcomq;
                                dgv.Rows[currentRow].Cells[i + 3].Value = Convert.ToString(lsumq);
                                dgv.Rows[compareRow].Cells[i + 3].Value = null;
                            }
                            dgv.Rows.Remove(dgv.Rows[compareRow]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show("Variables or Parameters you passed are not accurate or don't exist!", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        public void fillOrderNumber(DataGridView dgv)
        {
            for (int cRow = 0; cRow < (dgv.Rows.Count - 1); cRow++)
            {
                dgv.Rows[cRow].Cells[0].Value = cRow + 1;
            }
        }
    public List<int> FindRebarColumn(Worksheet wsh)
        {
            Excel.Range rng = wsh.UsedRange;
            int lcBOQ = rng.Columns.Count;
            List<int> rebarCols = new List<int>();
            //With / If structural file : find 3 rebar columns, add them into list<int> rebarCols
            //With / If architectural file : assign list<int> rebarCols = column 100 th
            for (int r = 1; r <= lcBOQ; r++)
            {
                string cellValue = wsh.Cells[1, r].Value2;
                if (cellValue != null)
                {
                    if (cellValue.Contains("<") == true | cellValue.Contains("=") == true | cellValue.Contains(">") == true)
                    {
                        rebarCols.Add(r);
                    }
                }
            }
            //If rebarCols.Count == 0 : Have NO Rebar columns in sheet (both Structural File and Architectural File)
            if (rebarCols.Count == 0)
            {
                rebarCols.Add(100);
                rebarCols.Add(100);
                rebarCols.Add(100);
            }
            return rebarCols;
        }
        public int FindSubtotalRow(Worksheet wsh, int refRow)
        {
            int nextRow = refRow;
            do
            {
                nextRow++;
            } while (Convert.ToString(wsh.Cells[nextRow, 1].Value2) != "Subtotal");
            return nextRow;
        }
        public int FindOrderCharacter(string source, string str)
        {
            int orderNum = 0;
            int leng = source.Length;
            if (source.Contains(str) == true)
            {
                for (int i = 0; i < leng; i++)
                {
                    string sourceCha = Convert.ToString(source[i]);
                    if (sourceCha == str)
                    {
                        orderNum = i;
                        break;
                    }
                }
                return orderNum;
            }
            else
            {
                return (leng - 1);
            }
        }

        private void bTest_Click(object sender, EventArgs e)
        {
            //int xuathientai = FindOrderCharacter("khoa (1993)", "(");
            //MessageBox.Show(" ( character appear at " + xuathientai + " character number");
        }

        private void cbFill_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbFill.Text == "Fill with a Selective Row")
            {bSelectRow.Enabled = true;}
            else
            {bSelectRow.Enabled = false;}
        }
    }
}
