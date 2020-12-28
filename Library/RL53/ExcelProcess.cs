using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laporan_Automation.Library.RL53
{
    class ExcelProcess
    {
        private static Dictionary<string, string> penyakit10besar { set; get; }
        private static Dictionary<string, Dictionary<string, int>> hidup { set; get; }

        public static void Process(Dictionary<string, OpenFileDialog> files)
        {
            ReadProcess(files["10Besar"], "10Besar"); //10 Besar wajib ditaruh pertama

            ReadProcess(files["Hidup"], "Hidup");

            WriteProcess(files["Laporan"]); //Write wajib terakhir
        }

        private static void ReadProcess(OpenFileDialog ofd, string name)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(ofd.FileName);
                string path = ofd.FileName;
                string new_path = Helper.ConvertXLS_XLSX(fileInfo);
                FileInfo new_fileinfo = new FileInfo(new_path);

                ExcelPackage excel = new ExcelPackage(new_fileinfo);
                ExcelWorksheet workSheet = excel.Workbook.Worksheets[0];


                switch (name)
                {

                    //------------------------------------- 10 BESAR PENYAKIT ----------------------------------- //
                    case "10Besar":
                        penyakit10besar = new Dictionary<string, string>();
                        for (int i = 8; i <= 17; i++)
                        {
                            penyakit10besar.Add(workSheet.Cells["C" + i].Value.ToString(), workSheet.Cells["E" + i].Value.ToString());
                        }
                        break;
                    // ------------------------------------------------------------------------------------------ //

                    //------------------------------------------- Hidup ----------------------------------------- //
                    case "Hidup":
                        hidup = new Dictionary<string, Dictionary<string, int>>()
                        {
                            ["laki-laki"] = new Dictionary<string, int>(),
                            ["perempuan"] = new Dictionary<string, int>(),
                        };
                        int row_count = workSheet.Dimension.End.Row;
                        foreach (var penyakit in penyakit10besar)
                        {
                            int counter_laki = 0;
                            int counter_perempuan = 0;
                            for (int i = 1; i <= row_count; i++)
                            {
                                if (workSheet.Cells["H" + i].Value == null) continue;
                                string jk = workSheet.Cells["H" + i].Value.ToString().ToLower();
                                string code = workSheet.Cells["N" + i].Value.ToString();
                                if (jk == "laki-laki" && code == penyakit.Key)
                                {
                                    counter_laki++;
                                }
                                else if (jk == "perempuan" && code == penyakit.Key)
                                {
                                    counter_perempuan++;
                                }
                            }
                            hidup["laki-laki"].Add(penyakit.Key, counter_laki);
                            hidup["perempuan"].Add(penyakit.Key, counter_perempuan);
                        }
                        break;
                    // ------------------------------------------------------------------------------------------ //

                    //----------------------------------------- Kematian ---------------------------------------- //
                    case "Kematian":
                        break;
                    // ------------------------------------------------------------------------------------------ //

                    default:
                        Console.WriteLine("Nothing to Process?!");
                        break;

                }

                File.Delete(path);
                Helper.ConvertXLSX_XLS(new_fileinfo);
                File.Delete(new_path);
            }
            catch (Exception er)
            {
                MessageBox.Show("Error :" + er.Message);
            }
        }

        private static void WriteProcess(OpenFileDialog ofd)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(ofd.FileName);
                string path = ofd.FileName;
                string new_path = Helper.ConvertXLS_XLSX(fileInfo);
                FileInfo new_fileinfo = new FileInfo(new_path);

                ExcelPackage excel = new ExcelPackage(new_fileinfo);
                ExcelWorksheet workSheet = excel.Workbook.Worksheets[0];


                //------------------------------------- 10 BESAR PENYAKIT ----------------------------------- //
                int index = 2;
                foreach (var penyakit in penyakit10besar)
                {
                    workSheet.Cells["H" + index].Value = penyakit.Key;
                    workSheet.Cells["H" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["I" + index].Value = penyakit.Value;
                    index++;
                }
                // ------------------------------------------------------------------------------------------ //

                //------------------------------------------- Hidup ----------------------------------------- //
                index = 2;
                foreach (var penyakit in penyakit10besar)
                {
                    foreach (var count in hidup["laki-laki"])
                    {
                        if (count.Key == penyakit.Key)
                        {
                            workSheet.Cells["J" + index].Value = count.Value;
                            workSheet.Cells["J" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            index++;
                        }
                    }
                }

                index = 2;
                foreach (var penyakit in penyakit10besar)
                {
                    foreach (var count in hidup["perempuan"])
                    {
                        if (count.Key == penyakit.Key)
                        {
                            workSheet.Cells["K" + index].Value = count.Value;
                            workSheet.Cells["K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            index++;
                        }
                    }
                }
                // ------------------------------------------------------------------------------------------ //

                //----------------------------------------- Kematian ---------------------------------------- //
                // ------------------------------------------------------------------------------------------ //

                excel.Save();

                File.Delete(path);
                Helper.ConvertXLSX_XLS(new_fileinfo);
                File.Delete(new_path);
            }
            catch (Exception er)
            {
                MessageBox.Show("Error :" + er.Message);
            }
        }
    }
}
