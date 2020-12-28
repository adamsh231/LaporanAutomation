using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laporan_Automation.Library
{
    class Helper
    {
        public static string ConvertXLS_XLSX(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }

        public static string ConvertXLSX_XLS(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsxFile = file.FullName;
            var wb = app.Workbooks.Open(xlsxFile);
            var xlsFile = xlsxFile.Substring(0, xlsxFile.Length - 1);
            wb.SaveAs(Filename: xlsFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlAddIn8);
            wb.Close();
            app.Quit();
            return xlsFile;
        }

        public static OpenFileDialog OpenFile(string filter = "Excel Files|*.xls;", bool multiselect = false)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = filter;
            ofd.Multiselect = multiselect;
            return ofd;
        }
    }
}
