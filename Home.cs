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
        }

        private void RL53_Click(object sender, EventArgs e)
        {
            RL53 open_rl53 = new RL53();
            open_rl53.Show();
        }
    }
}
