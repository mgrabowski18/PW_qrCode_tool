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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PW_qrCode_tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectTab(tabPage1);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void onTab1ChooseFile(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Plik Word (.docx ,.doc)|*.docx;*.doc";
            fileDialog.FileOk += delegate (object s, CancelEventArgs ev)
            {
                string ext = Path.GetExtension(fileDialog.FileName);
                if (ext != ".doc" && ext != ".docx")
                {
                    MessageBox.Show("Wybrany plik nie jest plikem Word!");
                    ev.Cancel = true;
                }
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fileDialog.FileName;
            }
        }

        private void onTab1ProcessFile(object sender, EventArgs e)
        {
            string path = textBox2.Text;
            string ext = Path.GetExtension(path);

            if (ext != ".doc" && ext != ".docx")
            {
                MessageBox.Show("Wybrany plik nie jest plikem Word!");
                return;
            }
            try
            {
                System.Diagnostics.Process.Start(path);
            }
            catch (Win32Exception exc)
            {
                MessageBox.Show(String.Format("Plik {0} nie istnieje!", path));
            }

            processWordFile(path);

        }

        protected void processWordFile(string path)
        {
            if (path.Length == 0)
            {
                return;
            }

        }
    }
}
