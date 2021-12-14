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
    public partial class FormSetupBOQ : Form
    {
        FormBOQ data;
        public FormSetupBOQ( FormBOQ data)
        {
            this.data = data;
            InitializeComponent();
        }

        private void FormSetupBOQ_Load(object sender, EventArgs e)
        {

        }

        private void bFirst_Click(object sender, EventArgs e)
        {

        }

        public void DoesSheetExists(string sht, Workbook wb)
        {

        }
    }
}
