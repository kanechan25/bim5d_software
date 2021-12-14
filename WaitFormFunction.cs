using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QSKSKS
{
    public class WaitFormFunction
    {
        WaitingForm wait;
        Thread loadThread;
        public void Show()
        {
            loadThread = new Thread(new ThreadStart(LoadingProcess));
            loadThread.Start();
        }
        public void Show(Form parent)
        {
            loadThread = new Thread(new ParameterizedThreadStart(LoadingProcess));
            loadThread.Start(parent);
        }
        public void Close()
        {
            if (wait != null)
            {
                wait.BeginInvoke(new System.Threading.ThreadStart(wait.CloseWaitForm));
                wait = null;
                loadThread = null;

            }
        }
        private void LoadingProcess()
        {
            wait = new WaitingForm();
            wait.ShowDialog();
        }
        private void LoadingProcess(object parent)
        {
            Form openParent = parent as Form;
            wait = new WaitingForm(openParent);
            wait.ShowDialog();
        }
    }
}
