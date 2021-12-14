using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QSKSKS
{
    public partial class WaitingForm : Form
    {
        public WaitingForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;
            this.TopMost = true;
        }
        public WaitingForm(Form parent)
        {
            InitializeComponent();
            if (parent != null)
            {
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new Point(parent.Location.X + parent.Width / 2 - this.Width / 2, parent.Location.Y + parent.Height / 2 - this.Height / 2);
                this.TopMost = true;
            }
            else
            {
                this.StartPosition = FormStartPosition.CenterParent;
                this.TopMost = true;
            }
        }
        private void WaitingForm_Load(object sender, EventArgs e)
        {

        }
        public void CloseWaitForm()
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
            if (true)
            {

            }
        }
    }
}
