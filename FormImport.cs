using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace QSKSKS
{
    public partial class FormTTD : Form
    {
        public FormTTD()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        public string[] ArrayPath = {"C:\\TTD BIM 5D\\txt\\ACatePath.txt", "C:\\TTD BIM 5D\\txt\\AExpPath.txt", "C:\\TTD BIM 5D\\txt\\SCatePath.txt",
            "C:\\TTD BIM 5D\\txt\\SExpPath.txt","C:\\TTD BIM 5D\\txt\\RebarPath.txt", "C:\\TTD BIM 5D\\txt\\Mec1Path.txt", "C:\\TTD BIM 5D\\txt\\Mec2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Mec3Path.txt","C:\\TTD BIM 5D\\txt\\Mec4Path.txt","C:\\TTD BIM 5D\\txt\\Ele1Path.txt", "C:\\TTD BIM 5D\\txt\\Ele2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Ele3Path.txt","C:\\TTD BIM 5D\\txt\\Ele4Path.txt", "C:\\TTD BIM 5D\\txt\\Plu1Path.txt","C:\\TTD BIM 5D\\txt\\Plu2Path.txt",
            "C:\\TTD BIM 5D\\txt\\Plu3Path.txt","C:\\TTD BIM 5D\\txt\\Plu4Path.txt","C:\\TTD BIM 5D\\txt\\FF1Path.txt","C:\\TTD BIM 5D\\txt\\FF2Path.txt",
            "C:\\TTD BIM 5D\\txt\\FF3Path.txt","C:\\TTD BIM 5D\\txt\\FF4Path.txt","C:\\TTD BIM 5D\\txt\\BOQPath.txt"};
        public void bClose_Click(object sender, EventArgs e)
        {
            //Khai bao bien Path cua tat ca cac Path
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] ArrayGetPath = {tbACatePath.Text,tbAExpPath.Text, tbSCatePath.Text, tbSExpPath.Text,tbRebarPath.Text,tbMec1Path.Text,tbMec2Path.Text,
            tbMec3Path.Text, tbMec4Path.Text,tbEle1Path.Text,tbEle2Path.Text,tbEle3Path.Text,tbEle4Path.Text,tbPlu1Path.Text,tbPlu2Path.Text,tbPlu3Path.Text,
            tbPlu4Path.Text,tbFF1Path.Text,tbFF2Path.Text,tbFF3Path.Text,tbFF4Path.Text,tbBOQPath.Text};
            //Luu cac Path vao file txt
            for (int i = 0; i < ArrayPath.Length; i++)
            {
                StreamWriter txt = new StreamWriter(ArrayPath[i]);
                txt.Write(string.Empty);
                txt.WriteLine(ArrayGetPath[i]);
                txt.Close();
            }
            this.Close();
        }
        public void SaveWhenLoadFile()
        {
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] ArrayGetPath = {tbACatePath.Text,tbAExpPath.Text, tbSCatePath.Text, tbSExpPath.Text,tbRebarPath.Text,tbMec1Path.Text,tbMec2Path.Text,
            tbMec3Path.Text, tbMec4Path.Text,tbEle1Path.Text,tbEle2Path.Text,tbEle3Path.Text,tbEle4Path.Text,tbPlu1Path.Text,tbPlu2Path.Text,tbPlu3Path.Text,
            tbPlu4Path.Text,tbFF1Path.Text,tbFF2Path.Text,tbFF3Path.Text,tbFF4Path.Text,tbBOQPath.Text};
            //Luu cac Path vao file txt
            for (int i = 0; i < ArrayPath.Length; i++)
            {
                StreamWriter txt = new StreamWriter(ArrayPath[i]);
                txt.Write(string.Empty);
                txt.WriteLine(ArrayGetPath[i]);
                txt.Close();
            }
        }
        private void FormTTD_Load(object sender, EventArgs e)
        {
            //this.ControlBox = false;
            #region //Khi open Form, lay toan bo Path da luu vao Textbox...Code o duoi region
                try
                {
                StreamReader txt1 = new StreamReader(ArrayPath[0]);
                tbACatePath.Text = txt1.ReadToEnd();
                txt1.Close();
                //
                StreamReader txt2 = new StreamReader(ArrayPath[1]);
                tbAExpPath.Text = txt2.ReadToEnd();
                txt2.Close();
                //
                StreamReader txt3 = new StreamReader(ArrayPath[2]);
                tbSCatePath.Text = txt3.ReadToEnd();
                txt3.Close();
                //
                StreamReader txt4 = new StreamReader(ArrayPath[3]);
                tbSExpPath.Text = txt4.ReadToEnd();
                txt4.Close();
                //
                StreamReader txt5 = new StreamReader(ArrayPath[4]);
                tbRebarPath.Text = txt5.ReadToEnd();
                txt5.Close();
                //
                StreamReader txt6 = new StreamReader(ArrayPath[5]);
                tbMec1Path.Text = txt6.ReadToEnd();
                txt6.Close();
                //
                StreamReader txt7 = new StreamReader(ArrayPath[6]);
                tbMec2Path.Text = txt7.ReadToEnd();
                txt7.Close();
                //
                StreamReader txt8 = new StreamReader(ArrayPath[7]);
                tbMec3Path.Text = txt8.ReadToEnd();
                txt8.Close();
                //
                StreamReader txt9 = new StreamReader(ArrayPath[8]);
                tbMec4Path.Text = txt9.ReadToEnd();
                txt9.Close();
                //
                StreamReader txt10 = new StreamReader(ArrayPath[9]);
                tbEle1Path.Text = txt10.ReadToEnd();
                txt10.Close();
                //
                StreamReader txt11 = new StreamReader(ArrayPath[10]);
                tbEle2Path.Text = txt11.ReadToEnd();
                txt11.Close();
                //
                StreamReader txt12 = new StreamReader(ArrayPath[11]);
                tbEle3Path.Text = txt12.ReadToEnd();
                txt12.Close();
                //
                StreamReader txt13 = new StreamReader(ArrayPath[12]);
                tbEle4Path.Text = txt13.ReadToEnd();
                txt13.Close();
                //
                StreamReader txt14 = new StreamReader(ArrayPath[13]);
                tbPlu1Path.Text = txt14.ReadToEnd();
                txt14.Close();
                //
                StreamReader txt15 = new StreamReader(ArrayPath[14]);
                tbPlu2Path.Text = txt15.ReadToEnd();
                txt15.Close();
                //
                StreamReader txt16 = new StreamReader(ArrayPath[15]);
                tbPlu3Path.Text = txt16.ReadToEnd();
                txt16.Close();
                //
                StreamReader txt17 = new StreamReader(ArrayPath[16]);
                tbPlu4Path.Text = txt17.ReadToEnd();
                txt17.Close();
                //
                StreamReader txt18 = new StreamReader(ArrayPath[17]);
                tbFF1Path.Text = txt18.ReadToEnd();
                txt18.Close();
                //
                StreamReader txt19 = new StreamReader(ArrayPath[18]);
                tbFF2Path.Text = txt19.ReadToEnd();
                txt19.Close();
                //
                StreamReader txt20 = new StreamReader(ArrayPath[19]);
                tbFF3Path.Text = txt20.ReadToEnd();
                txt20.Close();
                //
                StreamReader txt21 = new StreamReader(ArrayPath[20]);
                tbFF4Path.Text = txt21.ReadToEnd();
                txt21.Close();
                //
                StreamReader txt22 = new StreamReader(ArrayPath[21]);
                tbBOQPath.Text = txt22.ReadToEnd();
                txt22.Close();
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            #endregion
        }


        #region //Button de Get toan bo cac Path cua tat ca cac file Excel...Code o duoi region
        private void bGetACatePath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbACatePath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetAExpPath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbAExpPath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetSCatePath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbSCatePath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetSExpPath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbSExpPath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetRebarPath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbRebarPath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetBOQPath_Click(object sender, EventArgs e)
        {
            openPath.ShowDialog();
            tbBOQPath.Text = openPath.FileName;
            SaveWhenLoadFile();
        }
        private void bGetMec1Path_Click(object sender, EventArgs e)
        {
            openPath2.ShowDialog();
            string s = openPath2.FileName;
            System.IO.Stream str;
            int count = 0;
            List<string> listMultiPath = new List<string>();
            foreach (string file in openPath2.FileNames)
            {
                if ((str = openPath2.OpenFile()) != null)
                {
                    count++;
                    listMultiPath.Add(file);
                }
            }
            if (count > 4)
            {
                MessageBox.Show("The file number you have chose is " + count + "." + "\n" + "You can only select maximum to 4 files once.", 
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tbMec1Path.Text = "";
                tbMec2Path.Text = "";
                tbMec3Path.Text = "";
                tbMec4Path.Text = "";
            }
            else
            {
                switch (count)
                {
                    case 1:
                        tbMec1Path.Text = listMultiPath[0];
                        break;
                    case 2:
                        tbMec1Path.Text = listMultiPath[0];
                        tbMec2Path.Text = listMultiPath[1];
                        break;
                    case 3:
                        tbMec1Path.Text = listMultiPath[0];
                        tbMec2Path.Text = listMultiPath[1];
                        tbMec3Path.Text = listMultiPath[2];
                        break;
                    case 4:
                        tbMec1Path.Text = listMultiPath[0];
                        tbMec2Path.Text = listMultiPath[1];
                        tbMec3Path.Text = listMultiPath[2];
                        tbMec4Path.Text = listMultiPath[3];
                        break;
                }
            }
            SaveWhenLoadFile();
        }
        private void bGetEle1Path_Click(object sender, EventArgs e)
        {
            openPath2.ShowDialog();
            string s = openPath2.FileName;
            System.IO.Stream str;
            int count = 0;
            List<string> listMultiPath = new List<string>();
            foreach (string file in openPath2.FileNames)
            {
                if ((str = openPath2.OpenFile()) != null)
                {
                    count++;
                    listMultiPath.Add(file);
                }
            }
            if (count > 4)
            {
                MessageBox.Show("The file number you have chose is " + count + "." + "\n" + "You can only select maximum to 4 files once.",
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tbEle1Path.Text = "";
                tbEle2Path.Text = "";
                tbEle3Path.Text = "";
                tbEle4Path.Text = "";
            }
            else
            {
                switch (count)
                {
                    case 1:
                        tbEle1Path.Text = listMultiPath[0];
                        break;
                    case 2:
                        tbEle1Path.Text = listMultiPath[0];
                        tbEle2Path.Text = listMultiPath[1];
                        break;
                    case 3:
                        tbEle1Path.Text = listMultiPath[0];
                        tbEle2Path.Text = listMultiPath[1];
                        tbEle3Path.Text = listMultiPath[2];
                        break;
                    case 4:
                        tbEle1Path.Text = listMultiPath[0];
                        tbEle2Path.Text = listMultiPath[1];
                        tbEle3Path.Text = listMultiPath[2];
                        tbEle4Path.Text = listMultiPath[3];
                        break;
                }
            }
            SaveWhenLoadFile();
        }
        private void bGetPlu1Path_Click(object sender, EventArgs e)
        {
            openPath2.ShowDialog();
            string s = openPath2.FileName;
            System.IO.Stream str;
            int count = 0;
            List<string> listMultiPath = new List<string>();
            foreach (string file in openPath2.FileNames)
            {
                if ((str = openPath2.OpenFile()) != null)
                {
                    count++;
                    listMultiPath.Add(file);
                }
            }
            if (count > 4)
            {
                MessageBox.Show("The file number you have chose is " + count + "." + "\n" + "You can only sPluct maximum to 4 files once.",
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tbPlu1Path.Text = "";
                tbPlu2Path.Text = "";
                tbPlu3Path.Text = "";
                tbPlu4Path.Text = "";
            }
            else
            {
                switch (count)
                {
                    case 1:
                        tbPlu1Path.Text = listMultiPath[0];
                        break;
                    case 2:
                        tbPlu1Path.Text = listMultiPath[0];
                        tbPlu2Path.Text = listMultiPath[1];
                        break;
                    case 3:
                        tbPlu1Path.Text = listMultiPath[0];
                        tbPlu2Path.Text = listMultiPath[1];
                        tbPlu3Path.Text = listMultiPath[2];
                        break;
                    case 4:
                        tbPlu1Path.Text = listMultiPath[0];
                        tbPlu2Path.Text = listMultiPath[1];
                        tbPlu3Path.Text = listMultiPath[2];
                        tbPlu4Path.Text = listMultiPath[3];
                        break;
                }
            }
            SaveWhenLoadFile();
        }
        private void bGetFF1Path_Click(object sender, EventArgs e)
        {
            openPath2.ShowDialog();
            string s = openPath2.FileName;
            System.IO.Stream str;
            int count = 0;
            List<string> listMultiPath = new List<string>();
            foreach (string file in openPath2.FileNames)
            {
                if ((str = openPath2.OpenFile()) != null)
                {
                    count++;
                    listMultiPath.Add(file);
                }
            }
            if (count > 4)
            {
                MessageBox.Show("The file number you have chose is " + count + "." + "\n" + "You can only sFFct maximum to 4 files once.",
                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tbFF1Path.Text = "";
                tbFF2Path.Text = "";
                tbFF3Path.Text = "";
                tbFF4Path.Text = "";
            }
            else
            {
                switch (count)
                {
                    case 1:
                        tbFF1Path.Text = listMultiPath[0];
                        break;
                    case 2:
                        tbFF1Path.Text = listMultiPath[0];
                        tbFF2Path.Text = listMultiPath[1];
                        break;
                    case 3:
                        tbFF1Path.Text = listMultiPath[0];
                        tbFF2Path.Text = listMultiPath[1];
                        tbFF3Path.Text = listMultiPath[2];
                        break;
                    case 4:
                        tbFF1Path.Text = listMultiPath[0];
                        tbFF2Path.Text = listMultiPath[1];
                        tbFF3Path.Text = listMultiPath[2];
                        tbFF4Path.Text = listMultiPath[3];
                        break;
                }
            }
            SaveWhenLoadFile();
        }

        #endregion

        private void bClear_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you sure Clear all the paths?", "Notice!", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            switch (dr)
            {
                case DialogResult.Yes:
                    tbACatePath.Text = "";
                    tbAExpPath.Text = "";
                    tbSCatePath.Text = "";
                    tbSExpPath.Text = "";
                    tbRebarPath.Text = "";
                    tbBOQPath.Text = "";
                    tbMec1Path.Text = "";
                    tbMec2Path.Text = "";
                    tbMec3Path.Text = "";
                    tbMec4Path.Text = "";
                    tbEle1Path.Text = "";
                    tbEle2Path.Text = "";
                    tbEle3Path.Text = "";
                    tbEle4Path.Text = "";
                    tbPlu1Path.Text = "";
                    tbPlu2Path.Text = "";
                    tbPlu3Path.Text = "";
                    tbPlu4Path.Text = "";
                    tbFF1Path.Text = "";
                    tbFF2Path.Text = "";
                    tbFF3Path.Text = "";
                    tbFF4Path.Text = "";
                    break;
                case DialogResult.No:
                    break;
            }
            SaveWhenLoadFile();
        }
    }
}
