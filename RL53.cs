using Laporan_Automation.Library;
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
        private static OpenFileDialog ofd { get; set; }
        private static ExcelPackage excel { get; set; }
        private static ExcelWorksheet workSheet { get; set; }

        public RL53()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void Button_10Besar_Click(object sender, EventArgs e)
        {
            try
            {
                ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files|*.xls;"; //Jika multiple butuh detect file
                ofd.Multiselect = false;
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    TextBox_10Besar.Text = ofd.SafeFileName;

                    FileInfo fileInfo = new FileInfo(ofd.FileName);
                    string path = ofd.FileName;
                    string new_path = ExcelConverter.ConvertXLS_XLSX(fileInfo);
                    FileInfo new_fileinfo = new FileInfo(new_path);

                    excel = new ExcelPackage(new_fileinfo);
                    workSheet = excel.Workbook.Worksheets[0];
                    workSheet.Cells["D11"].Value = "SAMLEKOM MAMANGGG!!!!";
                    excel.Save();

                    File.Delete(path);
                    ExcelConverter.ConvertXLSX_XLS(new_fileinfo);
                    File.Delete(new_path);
                }

            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message);
            }
        }
    }
}
