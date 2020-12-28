using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laporan_Automation
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void RL53_Click(object sender, EventArgs e)
        {
            this.Hide();
            RL53 open_rl53 = new RL53();
            open_rl53.ShowDialog();
            this.Show();
        }
    }
}
