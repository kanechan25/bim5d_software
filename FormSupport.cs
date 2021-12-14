using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
    public partial class FormSupport : Form
    {
        public FormSupport()
        {
            InitializeComponent();
        }
        WaitFormFunction waitForm = new WaitFormFunction();

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
        public string[] ArrayPath = {"C:\\TTD BIM 5D\\txt\\ACatePath.txt", "C:\\TTD BIM 5D\\txt\\AExpPath.txt", "C:\\TTD BIM 5D\\txt\\SCatePath.txt",
            "C:\\TTD BIM 5D\\txt\\SExpPath.txt","C:\\TTD BIM 5D\\txt\\RebarPath.txt", "C:\\TTD BIM 5D\\txt\\Mec1Path.txt", "C:\\TTD BIM 5D\\txt\\Mec2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Mec3Path.txt","C:\\TTD BIM 5D\\txt\\Mec4Path.txt","C:\\TTD BIM 5D\\txt\\Ele1Path.txt", "C:\\TTD BIM 5D\\txt\\Ele2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Ele3Path.txt","C:\\TTD BIM 5D\\txt\\Ele4Path.txt", "C:\\TTD BIM 5D\\txt\\Plu1Path.txt","C:\\TTD BIM 5D\\txt\\Plu2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Plu3Path.txt","C:\\TTD BIM 5D\\txt\\Plu4Path.txt","C:\\TTD BIM 5D\\txt\\FF1Path.txt","C:\\TTD BIM 5D\\txt\\FF2Path.txt",
            "C:\\TTD BIM 5D\\txt\\FF3Path.txt","C:\\TTD BIM 5D\\txt\\FF4Path.txt","C:\\TTD BIM 5D\\txt\\BOQPath.txt"};
        #region Toàn bộ phần Get Category List và Get Element Type List
        public string valueIDCol_DIEQCS = "ID Revit";
        private void bFormatCate_Click(object sender, EventArgs e)
        //Format gồm 3 phần:
        //  1. Đổi tên sheet
        //  2. Autofit sheet
        //  3. Chèn thêm cột Level
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
            if (rbCheckStr.Checked == false && rbCheckArch.Checked == false)
            {
                MessageBox.Show("Please choosing a Circle Button that is Structure or Architecture.",
                "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    // Case 1: Structure File
                    if (rbCheckStr.Checked == true)
                    {
                        string myPathS = @"C:\TTD BIM 5D\txt\SCatePath.txt";
                        StreamReader txt = new StreamReader(myPathS);
                        string myFile = txt.ReadToEnd();
                        txt.Close();
                        var rootPath = Path.GetFullPath(myFile);
                        Workbook oWB = oXL.Workbooks.Open(rootPath);
                        ChangeSheetName(oWB);
                        AddLevelColumn(oWB);
                        FormatSheets(oWB);
                        AddRebarColumn(oWB);
                        //Save và Close app
                        oWB.Save();
                        oXL.Quit();
                        MessageBox.Show("Done!" + "\n" + "You have just formated: " + "\n" + " Change all sheet names,"
                            + "\n" + " Add a column that name \"Level\"," + "\n" + " Add 3 rebar columns.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        closeApplication(oXL);
                    }
                    // Case 2: Architect File
                    else if (rbCheckArch.Checked == true)
                    {
                        string myPathA = @"C:\TTD BIM 5D\txt\ACatePath.txt";
                        StreamReader txt = new StreamReader(myPathA);
                        string myFile = txt.ReadToEnd();
                        txt.Close();
                        var rootPath = Path.GetFullPath(myFile);
                        Workbook oWB = oXL2.Workbooks.Open(rootPath);
                        ChangeSheetName(oWB);
                        AddLevelColumn(oWB);
                        FormatSheets(oWB);
                        //Save và Close app
                        oWB.Save();
                        oXL2.Quit();
                        MessageBox.Show("Done!" + "\n" + "You have just formated: " + "\n" + " Change all sheet names,"
                            + "\n" + " Add a column that name \"Level\".", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        closeApplication(oXL2);
                    }
                }
                catch (Exception ex)
                {
                    closeApplication(oXL2);
                    closeApplication(oXL);
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        public void FormataSheet(Worksheet oSh)
        {
            Excel.Range usedRng = oSh.UsedRange;
            usedRng.ColumnWidth = 50;
            usedRng.Columns.AutoFit();
            usedRng.Rows.AutoFit();
        }
        public void FormatSheets(Workbook oWB)
        {
            oWB.Activate();
            int wsCount = oWB.Worksheets.Count;
            for (int s = 1; s <= wsCount; s++)
            {
                //Tìm lastRow và lastCol
                Excel.Worksheet oSheet = oWB.Worksheets[s];
                Excel.Range usedRng = oSheet.UsedRange;
                usedRng.ColumnWidth = 50;
                usedRng.Columns.AutoFit();
                usedRng.Rows.AutoFit();
            }
        }
        public void ChangeSheetName(Workbook oWB)
        {
            oWB.Activate();
            int wsCount = oWB.Worksheets.Count;
            for (int s = 1; s <= wsCount; s++)
            {
                //Tìm lastRow và lastCol
                Excel.Worksheet oSheet = oWB.Worksheets[s];
                Excel.Range usedRng = oSheet.UsedRange;
                Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = lastCell.Row;
                int lastCol = lastCell.Column;
                int elementCol = (FindColNumbyName(oWB, s, "Element Type"));
                oWB.Worksheets[s].Name = oSheet.Cells[2, elementCol].Value;
            }
        }
        public void AddRebarColumn(Workbook oWB)
        {
            oWB.Activate();
            int wsCount = oWB.Worksheets.Count;
            //B1. Duyệt lần lượt từng sheet
            for (int i = 1; i <= wsCount; i++)
                {
                    //Tìm lastRow và lastCol
                    Excel.Worksheet oSheet = oWB.Worksheets[i];
                    Excel.Range usedRng = oSheet.UsedRange;
                    //Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int lastRow = usedRng.Rows.Count;
                    int lastCol = usedRng.Columns.Count;
                    //B2. thêm các cột bên cạnh các cột KL Rebar
                    oSheet.Cells[1, lastCol + 1] = "d<=16";
                    oSheet.Cells[1, lastCol + 2] = "d<=25";
                    oSheet.Cells[1, lastCol + 3] = "d=29";
                }
            }
        public void AddLevelColumn(Workbook oWB)
        {
            oWB.Activate();
            int wsCount = oWB.Worksheets.Count;
            //B1. Duyệt lần lượt từng sheet
            for (int i = 1; i <= wsCount; i++)
            {
                //Tìm lastRow và lastCol
                Excel.Worksheet oSheet = oWB.Worksheets[i];
                Excel.Range usedRng = oSheet.UsedRange;
                Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = lastCell.Row;
                int lastCol = lastCell.Column;
                //Tìm các cột
                int col1 = (FindColNumbyChr(oWB, i, "("));
                int col2 = (FindColNumbyName(oWB, i, "Name"));
                int col3 = (FindColNumbyName(oWB, i, "Floor"));
                //MessageBox.Show(oWB.Worksheets[i].Name + " có cột ( là cột " + col1 + ", cột Name là cột " + col2 + ", cột Floor là cột " + col3);
                //B2. Insert thêm các cột bên cạnh các cột KL
                oSheet.Columns[col1].Insert();  //Insert 1 cột sang bên trái cột col1
                oSheet.Cells[1, col1] = "Level";

                //B3. Tìm row của các Level trong sheet theo các điều kiện
                for (int j = 2; j < lastRow; j++)
                {
                    var valueCellNameCol = oSheet.Cells[j, col2].Value;
                    bool contain1 = (Convert.ToString(valueCellNameCol)).Contains("Add");
                    bool contain2 = (Convert.ToString(valueCellNameCol)).Contains("Deduct");
                    if ((contain1 == false) && (contain2 == false))
                    {
                        if (oSheet.Cells[j, 1].Font.Bold == false)
                        {
                            oSheet.Cells[j, col1].Value = oSheet.Cells[j, col3].Value;
                        }
                    }
                }
            }
        }
        public bool IsNumeric(string input)
        {
            int number;
            return int.TryParse(input, out number);
        }
        public void CheckIsNumberic(string test)
        {
            if (IsNumeric(test) == true)
            {
                MessageBox.Show(test + " đúng là số");
            }
            else
            {
                MessageBox.Show(test + " không phải là số");
            }
        }
        public int FindColNumbyChr(Workbook oWB, int shNum, string Chr)
        //Tìm cột (số hiệu cột) đầu tiên từ trái qua phải chưa 1 ký tự nào đó (Chr)
        {
            /*myFile là text link path trong file.txt
            oXL.DisplayAlerts = false;
            oXL = new Excel.Application();
            var rootPath = Path.GetFullPath(myFile);
            oWB = oXL.Workbooks.Open(rootPath); */
            Worksheet oSh = oWB.Worksheets[shNum];
            string shName = oSh.Name;
            oSh.Activate();
            Excel.Range usedRange = oSh.UsedRange;
            Excel.Range lastCell = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = lastCell.Row;
            int lastCol = lastCell.Column;
            int i = 1;
            for (i = 1; i < lastCol; i++)
            {
                string stringCell = oSh.Cells[1, i].Value;
                if (stringCell == null)
                {
                    i++;
                }
                else
                {
                    int countAppear = CountChr(stringCell, Chr);
                    if (countAppear > 0)
                    {
                        break;
                    }
                    int Col = i;
                }
            }
            return i;
            //MessageBox.Show(lastRow + ", " + lastCol);
            //oXL.Visible = true;
            //oXL.DisplayAlerts = true;
        }
        public int FindColNumbyName(Workbook oWB, int shNum, string words)
        //Tìm cột (số hiệu cột) đầu tiên từ trái qua phải chưa 1 ký tự nào đó (Chr)
        {
            /*myFile là text link path trong file.txt
            oXL.DisplayAlerts = false;
            oXL = new Excel.Application();
            var rootPath = Path.GetFullPath(myFile);
            oWB = oXL.Workbooks.Open(rootPath);*/
            Worksheet oSh = oWB.Worksheets[shNum];
            string shName = oSh.Name;
            oSh.Activate();
            Excel.Range usedRange = oSh.UsedRange;
            Excel.Range lastCell = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = lastCell.Row;
            int lastCol = lastCell.Column;
            int i = 1;
            for (i = 1; i < lastCol; i++)
            {
                string stringCell = oSh.Cells[1, i].Value;
                if (stringCell == null)
                {
                    i++;
                }
                else
                {
                    bool containString = stringCell.Contains(words);
                    if (containString == true)
                    {
                        break;
                    }
                    int Col = i;
                }
            }
            return i;
            //MessageBox.Show(lastRow + ", " + lastCol);
            //oXL.Visible = true;
            //oXL.DisplayAlerts = true;
        }
        public int FindColNumChrXXX(string myFile, int shNum, string Chr)
        {   //Tìm cột (số hiệu cột) đầu tiên từ trái qua phải chưa 1 ký tự nào đó (Chr)
            //myFile là text link path trong file .txt
            //oXL.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            var rootPath = Path.GetFullPath(myFile);
            Excel.Workbook oWB = oXL.Workbooks.Open(rootPath);
            Excel.Worksheet oSh = (Excel.Worksheet)oWB.Worksheets[shNum];
            string shName = oSh.Name;
            oSh.Activate();
            Excel.Range usedRange = oSh.UsedRange;
            Excel.Range lastCell = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = lastCell.Row;
            int lastCol = lastCell.Column;
            int i = 1;
            for (i = 1; i < lastCol; i++)
            {
                string stringCell = oSh.Cells[1, i].Value;
                if (stringCell == null)
                {
                    i++;
                }
                else
                {
                    int countAppear = CountChr(stringCell, Chr);
                    if (countAppear > 0)
                    {
                        break;
                    }
                    int Col = i;
                }
            }
            closeApplication(oXL);
            return i;

            //MessageBox.Show(lastRow + ", " + lastCol);
            //oXL.Visible = true;
            //oXL.DisplayAlerts = true;
        }
        public int CountChr(string chuoiKyTu, string Chr)
        {
            int CountChr = 0;
            for (int i = 0; i < chuoiKyTu.Length; i++)
            {
                if (Convert.ToString(chuoiKyTu[i]) == Chr)
                {
                    CountChr++;
                }
            }
            return CountChr; ;
        }
        public void ReadWriteEdit(string fileName, int shNum)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.DisplayAlerts = false;
            //fileName : link path của Excel file
            //shNum : số hiệu của Sheet (1,2,3...)
            //1. Get all text in .txt file that contains path
            StreamReader txt = new StreamReader(fileName);
            string myFile = txt.ReadToEnd();
            txt.Close();
            object misvalue = System.Reflection.Missing.Value;
            //2. Read Excel File
            try
            {
                oXL = new Excel.Application();
                var rootPath = Path.GetFullPath(myFile);
                Workbook oWB = oXL.Workbooks.Open(rootPath);
                Worksheet oSh = (Excel.Worksheet)oWB.Worksheets[shNum];
                //oSh.Cells[1, 1] = tbText.Text;
                oWB.Save();
                oWB.Close();
                oXL.Quit();
                closeApplication(oXL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            oXL.DisplayAlerts = true;
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
        public void OpenFile(string xlpath, int Sh)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            string myPath = xlpath;
            if (System.IO.File.Exists(myPath))
            {
                //Get all text in .txt file that contains path
                StreamReader txt = new StreamReader(myPath);
                string myFile = txt.ReadToEnd();
                txt.Close();
                var rootPath = Path.GetFullPath(myFile);
                if (System.IO.File.Exists(myFile))
                {
                    //Open Excel File to View
                    Object misValue = System.Reflection.Missing.Value;
                    Workbook wb = oXL.Application.Workbooks.Open(rootPath, misValue, false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[Sh];
                    xlsh.Activate();
                    oXL.Visible = true;
                }
                else
                {
                    MessageBox.Show(" Unable to open Workbook because of uncorrect path." + "\n" + " Please check Import Form again!!",
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show(" Unable to open Workbook because of uncorrect path." + "\n" + " Please check Import Form again!!",
                "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            closeApplication(oXL);
        }
        private void bOK_Click(object sender, EventArgs e)
        {
            waitForm.Close();
            //Lưu các List vào file .txt
            if (rbCheckArch.Checked == true)
            {
                saveDGV(pathACate, dgvListCate);
                saveDGV(pathALevel, dgvListLevel);
                saveDGV(pathARowLevel, dgvRLL);
                saveDGV(pathAEle, dgvListEle);
                saveDGV(pathARowEle, dgvREL);
                saveDGV(pathAEleDep, dgvLDO);
                this.Close();
            }
            else if (rbCheckStr.Checked == true)
            {
                saveDGV(pathSCate, dgvListCate);
                saveDGV(pathSLevel, dgvListLevel);
                saveDGV(pathSRowLevel, dgvRLL);
                saveDGV(pathSEle, dgvListEle);
                saveDGV(pathSRowEle, dgvREL);
                saveDGV(pathSEleDep, dgvLDO);
                this.Close();
            }
        }
        public void SaveAfterLoad()
        {
            if (rbCheckArch.Checked == true)
            {
                saveDGV(pathACate, dgvListCate);
                saveDGV(pathALevel, dgvListLevel);
                saveDGV(pathARowLevel, dgvRLL);
                saveDGV(pathAEle, dgvListEle);
                saveDGV(pathARowEle, dgvREL);
                saveDGV(pathAEleDep, dgvLDO);
            }
            else if (rbCheckStr.Checked == true)
            {
                saveDGV(pathSCate, dgvListCate);
                saveDGV(pathSLevel, dgvListLevel);
                saveDGV(pathSRowLevel, dgvRLL);
                saveDGV(pathSEle, dgvListEle);
                saveDGV(pathSRowEle, dgvREL);
                saveDGV(pathSEleDep, dgvLDO);
            }
        }
        public void saveDGV(string myPathTXT, DataGridView dgv)
        {
            TextWriter txt = new StreamWriter(myPathTXT);
            for (int rc = 0; rc < dgv.Rows.Count; rc++)
            {
                txt.WriteLine(dgv.Rows[rc].Cells[0].Value.ToString());
            }
            txt.Close();
        }
        private void bImport_Click(object sender, EventArgs e)
        {
            FormTTD frm = new FormTTD();
            frm.Show();
        }

        private void FormSupport_Load(object sender, EventArgs e)
        {
            
            rbCheckStr.Checked = true;
            //bTest.Enabled = false;
            try
            {
                // read file .txt vào dataGridView
                GetDataToDGV(pathSCate, dgvListCate);
                GetDataToDGV(pathSLevel, dgvListLevel);
                GetDataToDGV(pathSRowLevel, dgvRLL);
                GetDataToDGV(pathSEle, dgvListEle);
                GetDataToDGV(pathSRowEle, dgvREL);
                GetDataToDGV(pathSEleDep, dgvLDO);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
        public void CreateTXT(string pathCreate)
        {
            FileStream fs = File.Create(pathCreate);
            fs.Close();
        }
        private void bTest_Click(object sender, EventArgs e)
        {
            #region Test Copy Paste sheet InputData
            /*string myPathS2 = @"C:\TTD BIM 5D\txt\SCatePath.txt";
            string myPathS = @"C:\TTD BIM 5D\txt\SExpPath.txt";
            //***File Category, copy Source, Sheet Cần chọn
            StreamReader txt2 = new StreamReader(myPathS2);
            string myFile2 = txt2.ReadToEnd();
            txt2.Close();
            var rootPath2 = Path.GetFullPath(myFile2);
            Workbook wbCateCopy = oXL2.Workbooks.Open(rootPath2);
            Worksheet wsCateCopy = wbCateCopy.Worksheets["Wall"];
            wsCateCopy.UsedRange.Copy(Type.Missing);
            //***File Expression, Sheet DIEQCS
            StreamReader txt = new StreamReader(myPathS);
            string myFile = txt.ReadToEnd();
            txt.Close();
            oXL = new Excel.Application();
            var rootPath = Path.GetFullPath(myFile);
            Workbook wbExpPaste = oXL.Workbooks.Open(rootPath);
            DoesSheetExists("InputData", wbExpPaste);
            Worksheet ShInputData = wbExpPaste.Worksheets["InputData"];
            object misValue = System.Reflection.Missing.Value;
            ShInputData.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            //oXL.DisplayAlerts = false;
            //oXL2.DisplayAlerts = false;
            wbCateCopy.Save();
            wbExpPaste.Save();
            wbCateCopy.Close();
            wbExpPaste.Close();
            //oXL.DisplayAlerts = true;
            //oXL2.DisplayAlerts = true;
            MessageBox.Show("Done! okokgoka", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            closeApplication();*/
            #endregion
            #region Test Find Column Number - NumColRange
            /*int selectedrowindex = dgvListCate.SelectedCells[0].RowIndex;
            DataGridViewRow selectedRow = dgvListCate.Rows[selectedrowindex];
            string CateValue = Convert.ToString(selectedRow.Cells[0].Value);
            string myPathS = @"C:\TTD BIM 5D\txt\SExpPath.txt";
            //string myPathS2 = @"C:\TTD BIM 5D\txt\SCatePath.txt";
            if (checkLinkPath(myPathS) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
            {
                waitForm.Show(this);
                1.1.File Category, copy Source, Sheet Cần chọn
            StreamReader txt2 = new StreamReader(myPathS2);
                string myFile2 = txt2.ReadToEnd();
                txt2.Close();
                var rootPath2 = Path.GetFullPath(myFile2);
                Workbook wbCateCopy = oXL2.Workbooks.Open(rootPath2);
                Worksheet wsCateCopy = wbCateCopy.Worksheets[CateValue];
                wsCateCopy.UsedRange.Copy(Type.Missing);
                //1.2.File Expression, Sheet DIEQCS
                StreamReader txt = new StreamReader(myPathS);
                string myFile = txt.ReadToEnd();
                txt.Close();
                oXL = new Excel.Application();
                var rootPath = Path.GetFullPath(myFile);
                Workbook wbExpPaste = oXL.Workbooks.Open(rootPath);
                DoesSheetExists("InputData", wbExpPaste);
                Worksheet ShInputData = wbExpPaste.Worksheets["InputData"];
                Worksheet QCS = wbExpPaste.Worksheets["DIEQCS"];
                object misValue = System.Reflection.Missing.Value;
                ShInputData.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll,
                Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                Tìm lastRow or lastCol của sheet InputData
            Excel.Range usedRng = ShInputData.UsedRange;
                Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                oXL.Visible = true;
                int lastRowInputData = lastCell.Row;
                int lastColInputData = lastCell.Column;
                int ColNum = NumColRange(wbExpPaste, "Level");
                int ColNum2 = NumColRange(wbExpPaste, "Name");
                MessageBox.Show(" Cot Level la " + ColNum);
                MessageBox.Show("Cot Name la " + ColNum2);
            }*/
            #endregion
            string value = "Volume = (0.400*3.720*0.200) = 0.298m3";
            string[] qtyExp = value.Split(' ');
            for (int i = 0; i < qtyExp.Length; i++)
            {
                MessageBox.Show("Value Header Quantity Expression is :" + qtyExp[i]);
            }
        }
        private void bGetCate_Click(object sender, EventArgs e)
        //1. List Category - bản chất là lấy list các Sheet của file Category
        //2. Format File Expression
        {
            dgvListCate.Rows.Clear();
            try
            {
                // Case 1: Structure File
                if (rbCheckStr.Checked == true && rbCheckArch.Checked == false)
                {
                    //1.1. Lấy List Category từ file Category
                    string myPathS = @"C:\TTD BIM 5D\txt\SCatePath.txt";
                    if (checkLinkPath(myPathS) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path đúng
                    {
                        Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                        StreamReader txt = new StreamReader(myPathS);
                        string myFile = txt.ReadToEnd();
                        txt.Close();
                        var rootPath = Path.GetFullPath(myFile);
                        Workbook oWB  = oXL.Workbooks.Open(rootPath);
                        GetListCategory(oWB);
                        //1.2. Save và Close app
                        oWB.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                        closeApplication(oXL);
                    }
                    else
                    { MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }
                // Case 2: Architecture File
                else if (rbCheckArch.Checked == true && rbCheckStr.Checked == false)
                {
                    string myPathA = @"C:\TTD BIM 5D\txt\ACatePath.txt";
                    if (checkLinkPath(myPathA) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                    {
                        Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                        StreamReader txt2 = new StreamReader(myPathA);
                        string myFile2 = txt2.ReadToEnd();
                        txt2.Close();
                        var rootPath2 = Path.GetFullPath(myFile2);
                        Workbook oWB2 = oXL2.Workbooks.Open(rootPath2);
                        GetListCategory(oWB2);
                        //Save và Close app
                        oWB2.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB2);
                        closeApplication(oXL2);
                    }
                    else
                    { MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            SaveAfterLoad();
        }
        List<string> listCate = new List<string>();
        List<string> listLevel = new List<string>();
        List<string> listRowLevel = new List<string>();
        List<string> listEle = new List<string>();
        List<string> listRowEle = new List<string>();
        List<string> listLevelDepend = new List<string>();
        public void GetListCategory(Workbook wb)
        {
            dgvListCate.Rows.Clear();
            listCate.Clear();
            wb.Activate();
            int wsCount = wb.Worksheets.Count;
            for (int s = 0; s < wsCount; s++)
            {
                listCate.Add(wb.Worksheets[s + 1].Name);
                dgvListCate.Rows.Add(listCate[s]);
            }
        }

        public void GetListLevel(Workbook oWB2)
        {
            try
            {
                oWB2.Activate();
                //Tìm lastRow và lastCol
                Excel.Worksheet oSheet = oWB2.Worksheets["DIEQCS"];
                Excel.Range usedRng = oSheet.UsedRange;
                Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = lastCell.Row;
                int lastCol = lastCell.Column;
                //1. Get listLevel to list
                int r = 1;
                do
                {
                    Range oRng = oSheet.Cells[r, 1];
                    string valueCell = Convert.ToString(oRng.Value2);
                    if (valueCell != null)
                    {
                        if (valueCell != "S/N" && valueCell.Contains(".") == false)
                        {
                            int isNum;
                            bool success = int.TryParse(valueCell, out isNum);
                            if (success == false)
                            {
                                listLevel.Add(valueCell);
                                string rowLevel = Convert.ToString(r);
                                listRowLevel.Add(rowLevel);
                            }
                        }
                    }
                    r++;
                } while (r <= lastRow);
                /*Method 2 : Using for looping:
                 for (int r = 1; r <= lastRow; r++)
                    {
                        Range oRng = oSheet.Cells[r, 1];
                        if (oSheet.Cells[r, 1].Font.Bold == true && Convert.ToString(oRng.Value2) != null)
                        {
                            try
                            {
                                if (oSheet.Cells[r, 1].Value != "S/N")
                                {
                                    if ((IsNumeric(oRng.Value)) == false)
                                    {
                                        MessageBox.Show("Danh sach listLevel lan luot la : " + oSheet.Cells[r, 1].Value + " tai dong so " + r);
                                        listLevel.Add(oSheet.Cells[r, 1].Value);
                                        string rowLevel = Convert.ToString(oSheet.Cells[r, 1].Row);
                                        listRowLevel.Add(rowLevel);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }*/
                listRowLevel.Add(Convert.ToString(lastRow+1));
                for (int l = 0; l < listLevel.Count; l++)
                {
                    dgvListLevel.Rows.Add(listLevel[l]);
                }
                for (int rl = 0; rl < listRowLevel.Count; rl++)
                {
                    dgvRLL.Rows.Add(listRowLevel[rl]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public void GetListElement(Workbook wb)
        {
            wb.Activate();
            //Tìm lastRow và lastCol
            Excel.Worksheet oSheet = wb.Worksheets["DIEQCS"];
            Excel.Range usedRng = oSheet.UsedRange;
            Excel.Range lastCell = usedRng.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = lastCell.Row;

            for (int r = 0; r < (listLevel.Count); r++)
            {
                try
                {
                    int top = Convert.ToInt32(listRowLevel[r]);
                    int bot = Convert.ToInt32(listRowLevel[r + 1]);
                    for (int rl = (top + 1); rl < bot; rl++)
                    {
                        Range oRng = oSheet.Cells[rl, 1];
                        string Element = Convert.ToString(oSheet.Cells[rl, 1].Value);
                        if (Element != null)
                        {
                            if (Element.Contains(".") == true)
                            {
                                //Add thêm Element vào List
                                string[] words = Element.Split('.');
                                listEle.Add(words[1]);
                                listLevelDepend.Add(listLevel[r]);
                                //Add thêm row number của Element vào List
                                string rowLevel = Convert.ToString(oRng.Row);
                                listRowEle.Add(rowLevel);
                                //MessageBox.Show("listEle add: " + words[1] + "\n" + "listLevelDepend add: " + listLevel[r] + "\n" + "listRowEle add: " + rowLevel);
                            }
                        }
                    }
                    //foreach (int rl in Enumerable.Range(top + 1, bot - 1))
                }
                catch (Exception)
                {
                }
            }
            listRowEle.Add(Convert.ToString(lastRow+1));
            //Lấy xong, điền vào các dgv
            for (int le = 0; le < listEle.Count; le++)
            {
                dgvListEle.Rows.Add(listEle[le]);
            }
            for (int rle = 0; rle < listRowEle.Count; rle++)
            {
                dgvREL.Rows.Add(listRowEle[rle]);
            }
            for (int rld = 0; rld < listLevelDepend.Count; rld++)
            {
                dgvLDO.Rows.Add(listLevelDepend[rld]);
            }
        }
        public void clearData()
        {
            dgvListEle.Rows.Clear();
            dgvListLevel.Rows.Clear();
            dgvREL.Rows.Clear();
            dgvRLL.Rows.Clear();
            dgvLDO.Rows.Clear();
        }
        public virtual bool checkLinkPath(string filePath)
        {
            StreamReader txt = new StreamReader(filePath);
            string myFile = txt.ReadToEnd();
            txt.Close();
            bool result = false;
            switch (myFile)
            {
                case null:
                    return result = false;
                case "":
                    return result = false;
            }
            if (System.IO.File.Exists(myFile))
            {
                return result = true;
            }
            return result;
        }
        private void bGetEleTypeList_Click(object sender, EventArgs e)
        //1. Lấy List Element và List Row Element trong file Expression
        //2. Lấy List Level và List Row Level trong file Expression
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            clearData();
            try
            {
                // Case 1: Structure File
                if (rbCheckStr.Checked == true)
                {
                    string myPathS = @"C:\TTD BIM 5D\txt\SExpPath.txt";
                    if (checkLinkPath(myPathS) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                    {
                        Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                        waitForm.Show(this);
                        StreamReader txt = new StreamReader(myPathS);
                        string myFile = txt.ReadToEnd();
                        txt.Close();
                        var rootPath = Path.GetFullPath(myFile);
                        Workbook oWB2 = oXL2.Workbooks.Open(rootPath);
                        Excel.Worksheet oSheet = oWB2.Worksheets["DIEQCS"];
                        oSheet.Cells[1, 5].Value = valueIDCol_DIEQCS;
                        oSheet.Cells[1, 5].Font.Bold = true;
                        //1.1. Lấy List Level và List Row Level trong file Expression
                        GetListLevel(oWB2);
                        //1.2. Lấy List Element và List Row Element trong file Expression
                        GetListElement(oWB2);
                        //1.3. Save và Close app
                        oWB2.Save();
                        oXL2.Quit();
                        waitForm.Close();
                        //MessageBox.Show("Done!", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        closeApplication(oXL2);
                        clearList();
                    }
                    else
                    {
                        MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                // Case 2: Architecture File
                else if (rbCheckArch.Checked == true)
                {
                    string myPathA = @"C:\TTD BIM 5D\txt\AExpPath.txt";
                    if (checkLinkPath(myPathA) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                    {
                        Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                        waitForm.Show(this);
                        StreamReader txt2 = new StreamReader(myPathA);
                        string myFile2 = txt2.ReadToEnd();
                        txt2.Close();
                        var rootPath2 = Path.GetFullPath(myFile2);
                        Workbook oWB2 = oXL.Workbooks.Open(rootPath2);
                        Excel.Worksheet oSheet = oWB2.Worksheets["DIEQCS"];
                        oSheet.Cells[1, 5].Value = valueIDCol_DIEQCS;
                        oSheet.Cells[1, 5].Font.Bold = true;
                        //1.1. Lấy List Level và List Row Level trong file Expression
                        GetListLevel(oWB2);
                        //1.2. Lấy List Element và List Row Element trong file Expression
                        GetListElement(oWB2);
                        //1.3. Save và Close app
                        oWB2.Save();
                        oXL.Quit();
                        waitForm.Close();
                        closeApplication(oXL);
                        clearList();
                    }
                    else
                    {
                        MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                stopwatch.Stop();
                TimeSpan ts = stopwatch.Elapsed;
                MessageBox.Show("Done! " + " Elapsed Time is: " + ts.Minutes + " Minutes", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            SaveAfterLoad();
        }
        public void clearList()
        {
            listCate.Clear();
            listLevel.Clear();
            listRowLevel.Clear();
            listEle.Clear();
            listRowEle.Clear();
            listLevelDepend.Clear();
        }
        private void rbCheckStr_CheckedChanged(object sender, EventArgs e)
        {
            dgvListCate.Rows.Clear();
            clearData();
            if (rbCheckArch.Checked == true)
            {
                GetDataToDGV(pathACate, dgvListCate);
                GetDataToDGV(pathALevel, dgvListLevel);
                GetDataToDGV(pathARowLevel, dgvRLL);
                GetDataToDGV(pathAEle, dgvListEle);
                GetDataToDGV(pathARowEle, dgvREL);
                GetDataToDGV(pathAEleDep, dgvLDO);
            }
            else if (rbCheckStr.Checked == true)
            {
                GetDataToDGV(pathSCate, dgvListCate);
                GetDataToDGV(pathSLevel, dgvListLevel);
                GetDataToDGV(pathSRowLevel, dgvRLL);
                GetDataToDGV(pathSEle, dgvListEle);
                GetDataToDGV(pathSRowEle, dgvREL);
                GetDataToDGV(pathSEleDep, dgvLDO);
            }
        }

        private void rbCheckArch_CheckedChanged(object sender, EventArgs e)
        {
            dgvListCate.Rows.Clear();
            clearData();
            if (rbCheckArch.Checked == true)
            {
                GetDataToDGV(pathACate, dgvListCate);
                GetDataToDGV(pathALevel, dgvListLevel);
                GetDataToDGV(pathARowLevel, dgvRLL);
                GetDataToDGV(pathAEle, dgvListEle);
                GetDataToDGV(pathARowEle, dgvREL);
                GetDataToDGV(pathAEleDep, dgvLDO);
            }
            else if (rbCheckStr.Checked == true)
            {
                GetDataToDGV(pathSCate, dgvListCate);
                GetDataToDGV(pathSLevel, dgvListLevel);
                GetDataToDGV(pathSRowLevel, dgvRLL);
                GetDataToDGV(pathSEle, dgvListEle);
                GetDataToDGV(pathSRowEle, dgvREL);
                GetDataToDGV(pathSEleDep, dgvLDO);
            }
        }

        private void clearListToolStripMenuItem_Click(object sender, EventArgs e)
        { dgvListCate.Rows.Clear(); }

        private void clearListToolStripMenuItem1_Click(object sender, EventArgs e)
        { dgvListLevel.Rows.Clear(); }

        public void DoesSheetExists(string sh, Workbook wb)
        {
            // Check xem có sheet sh hay không, nếu có clear Content  đi, nếu không có thì tạo sheet sh
            try
            {
                Excel.Worksheet wsh = wb.Worksheets[sh];
                Excel.Range usedRng = wsh.UsedRange;
                usedRng.ClearContents();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                //Create the worksheet
                Excel.Sheets worksheets = wb.Worksheets;
                var xlNewSheet = wb.Worksheets.Add(worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet.Name = sh;
            }
        }

        public void FindListIDColumn(Workbook wb, Worksheet wsh, int rc, string level)
        {

        }
        private void FormSupport_Closed(object sender, FormClosedEventArgs e)
        {
        }

        private void getIDForONECategoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bGetIDFor1_Click(sender, e);
        }
        #endregion

        #region Phần Get ID từ file Category vào file Expression
        private void bGetIDFor1_Click(object sender, EventArgs e)
        {
            //Trước khi chạy, hỏi người dùng chắc chắc chưa, ấn chạy sẽ không thể hoàn lại, và phải đợi
            if (dgvListCate.SelectedCells.Count > 0)
            {
                int selectedrowindex = dgvListCate.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dgvListCate.Rows[selectedrowindex];
                string CateValue = Convert.ToString(selectedRow.Cells[0].Value);
                string UpperValue = CateValue.ToUpper();
                DialogResult dialogResult = MessageBox.Show("Are you sure? " + "\n" + "You will start to Get ID for " + UpperValue + " category."
                    + "\n" + "You can't stop and it will take serveral minutes?" + "\n" + "Continue??", "Notice",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dialogResult == DialogResult.No)
                { }
                else if (dialogResult == DialogResult.Yes)
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    try
                    {
                        // Case 1: Structure File
                        if (rbCheckStr.Checked == true)
                        {
                            string myPathS = @"C:\TTD BIM 5D\txt\SExpPath.txt";
                            string myPathS2 = @"C:\TTD BIM 5D\txt\SCatePath.txt";
                            if (checkLinkPath(myPathS) == true && checkLinkPath(myPathS2) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                            {
                                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                                waitForm.Show(this);
                                //1.1.File Category, copy Source, Sheet Cần chọn
                                StreamReader txt2 = new StreamReader(myPathS2);
                                string myFile2 = txt2.ReadToEnd();
                                txt2.Close();
                                var rootPath2 = Path.GetFullPath(myFile2);
                                Workbook wbCateCopy = oXL2.Workbooks.Open(rootPath2);
                                Worksheet wsCateCopy = wbCateCopy.Worksheets[CateValue];
                                Excel.Range sourceRng = wsCateCopy.UsedRange;
                                sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();
                                //1.2.File Expression, Sheet DIEQCS
                                StreamReader txt = new StreamReader(myPathS);
                                string myFile = txt.ReadToEnd();
                                txt.Close();
                                var rootPath = Path.GetFullPath(myFile);
                                Workbook wbExpPaste = oXL.Workbooks.Open(rootPath);
                                DoesSheetExists("InputData", wbExpPaste);
                                Worksheet ShInputData = wbExpPaste.Worksheets["InputData"];
                                Worksheet QCS = wbExpPaste.Worksheets["DIEQCS"];
                                Excel.Range destnationRng = ShInputData.get_Range("A1");
                                destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                                //Sau khi copy xong sheet cần thiết (sheet CateValue) vào sheet InputData, tiến hành:
                                int lastInputDataRow = ShInputData.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                                //oXL.Visible = true;
                                //oXL2.DisplayAlerts = false;
                                wbCateCopy.Save();
                                wbCateCopy.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbCateCopy);
                                oXL2.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL2);
                                //1.3. Duyệt lần lượt theo từng Level, trong từng Level duyệt từng Category,
                                //khoanh vùng Range Category trùng với Category đang chỉ định
                                int numLevel = dgvListLevel.Rows.Count;
                                int colLevel = NumColRange(wbExpPaste, "InputData", "Level");
                                int colName = NumColRange(wbExpPaste, "InputData", "Name");
                                int colLocation = NumColRange(wbExpPaste, "InputData", "Location");
                                int colInputDataID = NumColRange(wbExpPaste, "InputData", "Summary Info");
                                int colIDExpression = NumColRange(wbExpPaste, "DIEQCS", valueIDCol_DIEQCS);
                                for (int il = 0; il < numLevel; il++) //Với mỗi Level trong file Expression, ta tiến hành:
                                {
                                    for (int ildo = 0; ildo < dgvLDO.Rows.Count; ildo++) //Xét với từng Level trong dgv.LDO
                                    {
                                        string levelRowLDO = Convert.ToString(dgvLDO.Rows[ildo].Cells[0].Value);
                                        string levelRowListLevel = Convert.ToString(dgvListLevel.Rows[il].Cells[0].Value);
                                        if (levelRowLDO == levelRowListLevel) //Nếu Level ở dgv.LDO trùng với Level ở dgv.LevelList
                                        {
                                            string listValue = Convert.ToString(dgvListEle.Rows[ildo].Cells[0].Value);
                                            if (listValue == CateValue) //Nếu dòng dgv.ListEle tương ứng bằng giá trị CateValue
                                            {
                                                int rowLevelTop = Convert.ToInt32(dgvREL.Rows[ildo].Cells[0].Value);
                                                int rowLevelBot = Convert.ToInt32(dgvREL.Rows[ildo + 1].Cells[0].Value);
                                                int topLevel = rowLevelTop + 1;
                                                int bottomLevel = rowLevelBot - 1;
                                                //Với mỗi dòng trong range của 1 Category trong 1 Level được khoanh vùng (trong sheet DIEQCS)
                                                for (int rl = topLevel; rl <= bottomLevel; rl++)
                                                {
                                                    for (int lripdt = 1; lripdt < lastInputDataRow + 1; lripdt++) //Xét từng dòng (trong sheet InputData) để tìm ID
                                                    {
                                                        string levelShInputData = ShInputData.Cells[lripdt, colLevel].Value2;
                                                        if (levelShInputData != null)
                                                        {
                                                            if (levelShInputData == Convert.ToString(dgvListLevel.Rows[il].Cells[0].Value)) //Xét duyệt Level
                                                            {
                                                                string nameShInputData = ShInputData.Cells[lripdt, colName].Value2;
                                                                string nameEleDIEQCS = QCS.Cells[rl, 2].Value2;
                                                                if (nameShInputData == nameEleDIEQCS) //Xét duyệt Name Element
                                                                {
                                                                    string locationBig = ShInputData.Cells[lripdt, colLocation].Value2;
                                                                    string locationSmall = QCS.Cells[rl, 3].Value2;
                                                                    if (locationBig.Contains(locationSmall) == true) //Xét duyệt Location
                                                                    {
                                                                        string valueIDShInputData = Convert.ToString(ShInputData.Cells[lripdt, colInputDataID].Value);
                                                                        QCS.Cells[rl, colIDExpression].Value = valueIDShInputData;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //oXL.Visible = true;
                                ShInputData.Delete();
                                //1.5. Save và Close app
                                wbExpPaste.Save();
                                wbExpPaste.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbExpPaste);
                                oXL.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                waitForm.Close();
                            }
                            else
                            {
                                MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        // Case 2: Architecture File
                        else if (rbCheckArch.Checked == true)
                        {
                            string myPathA = @"C:\TTD BIM 5D\txt\AExpPath.txt";
                            string myPathA2 = @"C:\TTD BIM 5D\txt\ACatePath.txt";
                            if (checkLinkPath(myPathA) == true && checkLinkPath(myPathA2) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                            {
                                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                                waitForm.Show(this);
                                //1.1.File Category, copy Source, Sheet Cần chọn
                                StreamReader txt2A = new StreamReader(myPathA2);
                                string myFile2 = txt2A.ReadToEnd();
                                txt2A.Close();
                                var rootPath2 = Path.GetFullPath(myFile2);
                                Workbook wbCateCopy = oXL2.Workbooks.Open(rootPath2);
                                Worksheet wsCateCopy = wbCateCopy.Worksheets[CateValue];
                                Excel.Range sourceRng = wsCateCopy.UsedRange;
                                sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();

                                //1.2.File Expression, Sheet DIEQCS
                                StreamReader txtA = new StreamReader(myPathA);
                                string myFile = txtA.ReadToEnd();
                                txtA.Close();
                                var rootPath = Path.GetFullPath(myFile);
                                Workbook wbExpPaste = oXL.Workbooks.Open(rootPath);
                                DoesSheetExists("InputData", wbExpPaste);
                                Worksheet ShInputData = wbExpPaste.Worksheets["InputData"];
                                Worksheet QCS = wbExpPaste.Worksheets["DIEQCS"];
                                Excel.Range destnationRng = ShInputData.get_Range("A1");
                                destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                                //Sau khi copy xong sheet cần thiết (sheet CateValue) vào sheet InputData, tiến hành:
                                int lastInputDataRow = ShInputData.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                                //oXL.Visible = true;
                                //oXL2.DisplayAlerts = false;
                                wbCateCopy.Save();
                                wbCateCopy.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbCateCopy);
                                oXL2.Quit();
                                closeApplication(oXL2);
                                //1.3. Duyệt lần lượt theo từng Level, trong từng Level duyệt từng Category,
                                //khoanh vùng Range Category trùng với Category đang chỉ định
                                int numLevel = dgvListLevel.Rows.Count;
                                int colLevel = NumColRange(wbExpPaste, "InputData", "Level");
                                int colName = NumColRange(wbExpPaste, "InputData", "Name");
                                int colLocation = NumColRange(wbExpPaste, "InputData", "Location");
                                int colInputDataID = NumColRange(wbExpPaste, "InputData", "Summary Info");
                                int colIDExpression = NumColRange(wbExpPaste, "DIEQCS", valueIDCol_DIEQCS);
                                //MessageBox.Show(" Cot Level la " + colLevel);
                                //MessageBox.Show("Cot Name la " + colName);
                                for (int il = 0; il < numLevel; il++) //Với mỗi Level trong file Expression, ta tiến hành:
                                {
                                    for (int ildo = 0; ildo < dgvLDO.Rows.Count; ildo++) //Xét với từng Level trong dgv.LDO
                                    {
                                        string levelRowLDO = Convert.ToString(dgvLDO.Rows[ildo].Cells[0].Value);
                                        string levelRowListLevel = Convert.ToString(dgvListLevel.Rows[il].Cells[0].Value);
                                        if (levelRowLDO == levelRowListLevel) //Nếu Level ở dgv.LDO trùng với Level ở dgv.LevelList
                                        {
                                            string listValue = Convert.ToString(dgvListEle.Rows[ildo].Cells[0].Value);
                                            if (listValue == CateValue) //Nếu dòng dgv.ListEle tương ứng bằng giá trị CateValue
                                            {
                                                int rowLevelTop = Convert.ToInt32(dgvREL.Rows[ildo].Cells[0].Value);
                                                int rowLevelBot = Convert.ToInt32(dgvREL.Rows[ildo + 1].Cells[0].Value);
                                                int topLevel = rowLevelTop + 1;
                                                int bottomLevel = rowLevelBot - 1;
                                                //Với mỗi dòng trong range của 1 Category trong 1 Level được khoanh vùng (trong sheet DIEQCS)
                                                for (int rl = topLevel; rl <= bottomLevel; rl++)
                                                {
                                                    for (int lripdt = 1; lripdt < lastInputDataRow + 1; lripdt++) //Xét từng dòng (trong sheet InputData) để tìm ID
                                                    {
                                                        string levelShInputData = ShInputData.Cells[lripdt, colLevel].Value2;
                                                        if (levelShInputData != null)
                                                        {
                                                            if (levelShInputData == Convert.ToString(dgvListLevel.Rows[il].Cells[0].Value)) //Xét duyệt Level
                                                            {
                                                                string nameShInputData = ShInputData.Cells[lripdt, colName].Value2;
                                                                string nameEleDIEQCS = QCS.Cells[rl, 2].Value2;
                                                                if (nameShInputData == nameEleDIEQCS) //Xét duyệt Name Element
                                                                {
                                                                    string locationBig = ShInputData.Cells[lripdt, colLocation].Value2;
                                                                    string locationSmall = QCS.Cells[rl, 3].Value2;
                                                                    if (locationBig.Contains(locationSmall) == true) //Xét duyệt Location
                                                                    {
                                                                        string valueIDShInputData = Convert.ToString(ShInputData.Cells[lripdt, colInputDataID].Value);
                                                                        QCS.Cells[rl, colIDExpression].Value = valueIDShInputData;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //oXL.DisplayAlerts = false;
                                ShInputData.Delete();
                                //1.5. Save và Close app
                                wbExpPaste.Save();
                                wbExpPaste.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbExpPaste);
                                oXL.Quit();
                                closeApplication(oXL);
                                waitForm.Close();
                            }
                            else
                            {
                                MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        stopwatch.Stop();
                        TimeSpan ts = stopwatch.Elapsed;
                        this.TopMost = true;
                        MessageBox.Show("Done! " + " Elapsed Time is: " + ts.Minutes + " Minutes", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information,MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        public int NumColRange(Workbook wb, string wsheet, string colName)
        {
            int colNum = 1;
            Worksheet wsh = wb.Worksheets[wsheet];
            for (int iCol = 1; iCol < 20; iCol++)
            {
                string cellValue = wsh.Cells[1, iCol].Value2;
                if (cellValue == colName)
                {
                    colNum = iCol;
                    return colNum;
                }
            }
            return colNum;
        }
        public int GetLastRow(Workbook wkb, string sheet, string column)
        {
            Microsoft.Office.Interop.Excel.Worksheet sht = wkb.Worksheets[sheet] as Worksheet;
            Microsoft.Office.Interop.Excel.Range range = sht.Range[column + ":" + column];
            range = range.Cells[range.Rows.Count, range.Column] as Range;
            return range.End[XlDirection.xlUp].Row;
        }
        public int GetLastCol(Workbook wkb, string sheet)
        {
            Microsoft.Office.Interop.Excel.Worksheet sht = wkb.Worksheets[sheet] as Worksheet;
            Excel.Range xlRange = sht.UsedRange;
            int colCount = xlRange.Columns.Count;
            return colCount;
        }
        private void bGetIDAll_Click(object sender, EventArgs e)
        {
            int countCate = dgvListCate.Rows.Count;
            //MessageBox.Show("Số Category là : " + countCate);
            for (int i = 0; i < countCate; i++)
            {
                dgvListCate.Rows[i].Selected = true;
                //bGetIDFor1_Click(sender, e);
            }
        }

        #endregion

        private void bGetExp_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure? " + "\n" + "You can't stop and it will take serveral minutes?" 
                                                        + "\n" + "Continue??", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.No)
            { }
            else if (dialogResult == DialogResult.Yes)
            {
                //You must run bGetEleTypeList Button before use this button
                if (dgvListLevel.Rows.Count > 0 && dgvRLL.Rows.Count > 0)
                {
                    //1. Open 2 files : Category and Expression, copy sheet DIEQCS to Category file, then close file Expression (Just handle only file)
                    try
                    {
                        Stopwatch stopwatch = new Stopwatch();
                        stopwatch.Start();
                        // Case 1: Structure File
                        if (rbCheckStr.Checked == true)
                        {
                            string myPathSExp = @"C:\TTD BIM 5D\txt\SExpPath.txt";
                            string myPathSCate = @"C:\TTD BIM 5D\txt\SCatePath.txt";
                            if (checkLinkPath(myPathSExp) == true && checkLinkPath(myPathSCate) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                            {
                                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                                waitForm.Show(this);
                                //1.1.File Expression, Sheet DIEQCS, copy Source
                                StreamReader txt2 = new StreamReader(myPathSExp);
                                string myFileExp = txt2.ReadToEnd();
                                txt2.Close();
                                var rootPathExp = Path.GetFullPath(myFileExp);
                                Workbook wbExpCopy = oXL2.Workbooks.Open(rootPathExp);
                                Worksheet QCS = wbExpCopy.Worksheets["DIEQCS"];
                                //Worksheet wsCateCopy = wbCateCopy.Worksheets[111111];
                                Excel.Range sourceRng = QCS.UsedRange;
                                sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();
                                //1.2.File Category, paste sheet DIEQCS to here
                                StreamReader txt = new StreamReader(myPathSCate);
                                string myFileCate = txt.ReadToEnd();
                                txt.Close();
                                var rootPath = Path.GetFullPath(myFileCate);
                                Workbook wbCatePaste = oXL.Workbooks.Open(rootPath);
                                DoesSheetExists("DIEQCS", wbCatePaste);
                                Worksheet DIEQCS = wbCatePaste.Worksheets["DIEQCS"];
                                Excel.Range destnationRng = DIEQCS.get_Range("A1");
                                destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                                FormataSheet(DIEQCS);
                                //destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths);
                                //Sau khi copy xong sheet cần thiết (sheet CateValue) vào sheet InputData, tiến hành tắt file Expression đi:
                                wbExpCopy.Save();
                                wbExpCopy.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbExpCopy);
                                oXL2.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL2);
                                //Bắt đầu lấy Expression từ sheet DIEQCS sang các sheet category khác
                                    FindExpressiontoCategory(wbCatePaste, DIEQCS); //***************************************************
                                DIEQCS.Delete();
                                //1.5. Save và Close app
                                wbCatePaste.Save();
                                wbCatePaste.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbCatePaste);
                                oXL.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                waitForm.Close();
                            }
                            else
                            {
                                MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        // Case 2: Architecture File
                        else if (rbCheckArch.Checked == true)
                        {
                            string myPathAExp = @"C:\TTD BIM 5D\txt\AExpPath.txt";
                            string myPathACate = @"C:\TTD BIM 5D\txt\ACatePath.txt";
                            if (checkLinkPath(myPathAExp) == true && checkLinkPath(myPathACate) == true) //Nếu trong file .txt có/không có gì hoặc Link Path SAI - tức là đã/chưa Import Path
                            {
                                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Application oXL2 = new Microsoft.Office.Interop.Excel.Application();
                                waitForm.Show(this);
                                //1.1.File Expression, Sheet DIEQCS, copy Source
                                StreamReader txt2 = new StreamReader(myPathAExp);
                                string myFileExp = txt2.ReadToEnd();
                                txt2.Close();
                                var rootPathExp = Path.GetFullPath(myFileExp);
                                Workbook wbExpCopy = oXL2.Workbooks.Open(rootPathExp);
                                Worksheet QCS = wbExpCopy.Worksheets["DIEQCS"];
                                //Worksheet wsCateCopy = wbCateCopy.Worksheets[111111];
                                Excel.Range sourceRng = QCS.UsedRange;
                                sourceRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).Copy();
                                //1.2.File Category, paste sheet DIEQCS to here
                                StreamReader txt = new StreamReader(myPathACate);
                                string myFileCate = txt.ReadToEnd();
                                txt.Close();
                                var rootPath = Path.GetFullPath(myFileCate);
                                Workbook wbCatePaste = oXL.Workbooks.Open(rootPath);
                                DoesSheetExists("DIEQCS", wbCatePaste);
                                Worksheet DIEQCS = wbCatePaste.Worksheets["DIEQCS"];
                                Excel.Range destnationRng = DIEQCS.get_Range("A1");
                                destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                                FormataSheet(DIEQCS);
                                //destnationRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths);
                                //Sau khi copy xong sheet cần thiết (sheet CateValue) vào sheet InputData, tiến hành tắt file Expression đi:
                                wbExpCopy.Save();
                                wbExpCopy.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbExpCopy);
                                oXL2.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL2);
                                //Bắt đầu lấy Expression từ sheet DIEQCS sang các sheet category khác
                                    FindExpressiontoCategory(wbCatePaste, DIEQCS); //***********************************************************
                                DIEQCS.Delete();
                                //1.5. Save và Close app
                                wbCatePaste.Save();
                                wbCatePaste.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbCatePaste);
                                oXL.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                waitForm.Close();
                            }
                            else
                            {
                                MessageBox.Show("You have not imported link path", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        stopwatch.Stop();
                        TimeSpan ts = stopwatch.Elapsed;
                        this.TopMost = true;
                        MessageBox.Show("Done! " + " Elapsed Time is: " + ts.Minutes + " Minutes", "Notice", MessageBoxButtons.OK,
                            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        public void FindExpressiontoCategory(Workbook wbk, Worksheet wshQCS)
        {
            int wsCount = wbk.Worksheets.Count;
            for (int i = 1; i <= wsCount; i++)
            {
                Worksheet wshCate = wbk.Worksheets[i];
                //If it's NOT DIEQCS sheet, we'll do it continue (do NOT anything with DIEQCS sheet)
                if (wshCate.Name != wshQCS.Name)
                {   //cumulative consideration sheet by sheet of categories
                    InsertExpressionCol(wshCate);
                    FindExpressionOneSheet(wbk, wshCate, wshQCS);
                }
            }
        }
        public void FindExpressionOneSheet(Workbook wbk, Worksheet wshCate, Worksheet wshQCS)
        {
            int QCS_IDRevit = NumColRange(wbk, wshQCS.Name, valueIDCol_DIEQCS);
            int QCS_Location = 3;
            int QCSQtyExp = 4;
            int Cate_ID = NumColRange(wbk, wshCate.Name, "Summary Info");
            int Cate_Location = NumColRange(wbk, wshCate.Name, "Location");
            int Cate_Level= NumColRange(wbk, wshCate.Name, "Level");
            Excel.Range usedRng = wshCate.UsedRange;
            int lrCate = usedRng.Rows.Count;
            int lcCate = usedRng.Columns.Count;
            string Cate_Type = wshCate.Name;
            for (int e = 0; e < dgvListEle.Rows.Count; e++)
            {
                string elementName = Convert.ToString(dgvListEle.Rows[e].Cells[0].Value);
                if (elementName == Cate_Type)
                {   
                    int upRow = Convert.ToInt32(dgvREL.Rows[e].Cells[0].Value);
                    int belowRow = Convert.ToInt32(dgvREL.Rows[e+1].Cells[0].Value);
                    for (int r = upRow+1; r < belowRow; r++)
                    {
                        string QCS_valueIDRevit = Convert.ToString(wshQCS.Cells[r, QCS_IDRevit].Value);
                        if (QCS_valueIDRevit != null)
                        {
                            for (int c = 2; c <= lrCate; c++)
                            {
                                string Cate_valueLevel = Convert.ToString(wshCate.Cells[c, Cate_Level].Value);
                                string Cate_valueID = Convert.ToString(wshCate.Cells[c, Cate_ID].Value);

                                if (Cate_valueLevel != null && Cate_valueID == QCS_valueIDRevit) //Check Element Type and ID are the same each other.
                                {
                                    string locationBig = Convert.ToString(wshCate.Cells[c, Cate_Location].Value);
                                    string locationSmall = Convert.ToString(wshQCS.Cells[r, QCS_Location].Value);
                                    if (locationBig.Contains(locationSmall) == true) //Check Location, If it's true, take corresponding Expression
                                    {   //Same location, same ID, now check Quantity Expression (maybe Volumn / Area...)
                                        int Cate_Qty1 = SpecialColNum(wshCate, "(", 1);
                                        for (int n = Cate_Qty1; n < (lcCate + 1); n++)
                                        {
                                            string Cate_CellQty = wshCate.Cells[1, n].Value;
                                            string QCS_CellQty = wshQCS.Cells[r, QCSQtyExp].Value;
                                            if (Cate_CellQty != null)
                                            {
                                                if (Cate_CellQty.Contains("(") == true)
                                                {
                                                    string[] qty = Cate_CellQty.Split('(');
                                                    string Cate_valueQty = qty[0];
                                                    string[] qtyExp = QCS_CellQty.Split(' ');
                                                    string QCS_valueQty = qtyExp[0];
                                                    if (Cate_valueQty == QCS_valueQty)
                                                    {
                                                        wshCate.Cells[c, n + 1].Value = QCS_CellQty;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        public int CountAppearTime(Excel.Range rng)
        {
            int count = 0;
            foreach (Excel.Range cell in rng)
            {
                string valueCell = cell.Value;
                if (valueCell.Contains("(") == true)
                {
                    count++;
                }
            }
            return count;
        }
        public void InsertExpressionCol(Worksheet wsh)
        {
            //Check quantity Columns (Volumn(m3), Area(m2)...), if the next empty column of them is existed, pass over, if not, insert column
            Excel.Range usedRng = wsh.UsedRange;
            int lrCate = usedRng.Rows.Count;
            int lcCate = usedRng.Columns.Count;
            //Excel.Range newRng = wsh.get_Range(wsh.Cells[1, 1], wsh.Cells[1, lcCate]);
            //int appearTime = CountAppearTime(newRng);
            int colQty1 = SpecialColNum(wsh, "(", 1);
            int numInsertCol = lcCate - colQty1;
            int col = colQty1;
            do
            {
                string valueCell = wsh.Cells[1, col].Value2;
                if (valueCell != null)
                {
                    if (valueCell.Contains("(") == true)
                    {
                        string cellNext = wsh.Cells[1, col+1].Value2;
                        if (cellNext == null)
                        {
                            //The Next Column is Column that need to insert, so break here (do NOTHING)
                        }
                        else
                        {
                            Range rng = (Excel.Range)wsh.Cells[1, col + 1];
                            rng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        }
                    }
                }
                col++;
            } while (col <= (lcCate + numInsertCol));
        }
        public int SpecialColNum(Worksheet wsh, string str, int orderNum)
        //Find the position (order Number) column letter that contain a special character
        // that appeared for the n time (calculate on cell by cell value), If NOT, return it's column 100th
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
    }
}

