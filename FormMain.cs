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

namespace QSKSKS
{
    public delegate string thePaths();
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            //this.TopMost = true;
        }

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
        #region All Code Here
        private void FormMain_Load(object sender, EventArgs e)
        {
            //Kiểm tra xem có thư mục txt hay không?
            string path = @"C:\TTD BIM 5D\txt";
            string fileCheck = @"C:\TTD BIM 5D\txt\ListACate.txt";
            string fileCheck2 = @"C:\TTD BIM 5D\txt\SCatePath.txt";
            if (Directory.Exists(path))
            {
                //Nếu tồn tại path rồi thì check tiếp có tồn tại file .txt hay không?
                if (File.Exists(fileCheck) && File.Exists(fileCheck2))
                {
                    //Nếu tồn tại file .txt rồi thì thôi, không thì else và tạo
                }
                else
                {
                    CreateTXT(pathACate);
                    CreateTXT(pathALevel);
                    CreateTXT(pathARowLevel);
                    CreateTXT(pathAEle);
                    CreateTXT(pathARowEle);
                    CreateTXT(pathAEleDep);

                    CreateTXT(pathSCate);
                    CreateTXT(pathSLevel);
                    CreateTXT(pathSRowLevel);
                    CreateTXT(pathSEle);
                    CreateTXT(pathSRowEle);
                    CreateTXT(pathSEleDep);

                    for (int i = 0; i < ArrayPath.Length; i++)
                    {
                        FileStream fs = File.Create(ArrayPath[i]);
                        fs.Close();
                    }
                }
            }
            else
            {
                Directory.CreateDirectory(path);
                CreateTXT(pathACate);
                CreateTXT(pathALevel);
                CreateTXT(pathARowLevel);
                CreateTXT(pathAEle);
                CreateTXT(pathARowEle);
                CreateTXT(pathAEleDep);

                CreateTXT(pathSCate);
                CreateTXT(pathSLevel);
                CreateTXT(pathSRowLevel);
                CreateTXT(pathSEle);
                CreateTXT(pathSRowEle);
                CreateTXT(pathSEleDep);

                for (int i = 0; i < ArrayPath.Length; i++)
                {
                    FileStream fs = File.Create(ArrayPath[i]);
                    fs.Close();
                }
            }
        }
        public void CreateTXT(string pathCreate)
        {
            FileStream fs = File.Create(pathCreate);
            fs.Close();
        }
        private void bImport_Click(object sender, EventArgs e)
        {
            FormTTD frm = new FormTTD();
            frm.Show();
            //frm.TopMost = true;
        }
        private void bSupport_Click(object sender, EventArgs e)
        {
            FormSupport frm = new FormSupport();
            frm.Show();
            //frm.TopMost = true;
        }

        private void bBOQ_Click(object sender, EventArgs e)
        {
            FormBOQ frm = new FormBOQ();
            frm.Show();
            //frm.TopMost = true;
        }

        private void bRebar_Click(object sender, EventArgs e)
        {
            FormRebar frmR = new FormRebar();
            frmR.Show();
        }
        private void FormMain_Closed(object sender, FormClosedEventArgs e)
        {

        }
        #endregion

        private void bClearProj_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure? " + "\n" + "All Data in the project will be deleted!"
                                                                                        + "\n" + "You won't undo or interrupt the deletion process!" + "\n" + "Continue??", "Notice",
                                                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.No)
            { }
            else if (dialogResult == DialogResult.Yes)
            {
                string[] listFiles = Directory.GetFiles(@"C:\TTD BIM 5D\txt\", "*.txt");
                foreach (string file in listFiles)
                {
                    var rootPath = Path.GetFullPath(file);
                    File.WriteAllText(rootPath, "");
                }
            }
        }
    }

    public static class pathList
        {
        public static string pathKKKK()
        {
            string str = "C:\\TTD BIM 5D\\txt\\ListACate.txt";
            return str;
        }

    }
}
