using Laporan_Automation.Library;
using Laporan_Automation.Library.RL53;
using OfficeOpenXml;
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

namespace Laporan_Automation
{
    public partial class RL53 : Form
    {
        public RL53()
        {
            InitializeComponent();
        }

        private void Button_10Besar_Click(object sender, EventArgs e)
        {
            TextBox_10Besar.Text = MainRL53.OpenAndSave("10Besar");
        }

        private void Button_Kematian_Click(object sender, EventArgs e)
        {
            TextBox_Kematian.Text = MainRL53.OpenAndSave("Kematian");
        }

        private void Button_Hidup_Click(object sender, EventArgs e)
        {
            TextBox_Hidup.Text = MainRL53.OpenAndSave("Hidup");
        }

        private void Button_Laporan_Click(object sender, EventArgs e)
        {
            TextBox_Laporan.Text = MainRL53.OpenAndSave("Laporan");
        }

        private async void Button_Process_Click(object sender, EventArgs e)
        {
            if (MainRL53.CheckFiles())
            {
                DialogResult result = MessageBox.Show("Apa file sudah benar ?", "Process", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result.Equals(DialogResult.OK))
                {
                    LoadingState(true);
                    await Task.Run(() => MainRL53.Process());
                    MainRL53.EmptyFilesDictionary();
                    MessageBox.Show("Process Complete!");
                    LoadingState(false);
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Please complete file first, before processing!");
            }
        }

        private void LoadingState(bool isLoading)
        {
            GB_Input.Enabled = !isLoading;
            GB_Output.Enabled = !isLoading;
            Button_Process.Enabled = !isLoading;
            Loading.Visible = isLoading;
        }

        private void RL53_FormClosing(object sender, FormClosingEventArgs e)
        {
            MainRL53.EmptyFilesDictionary();
        }
    }
}
