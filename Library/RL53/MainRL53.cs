using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laporan_Automation.Library.RL53
{
    class MainRL53
    {
        private static Dictionary<string, OpenFileDialog> files = new Dictionary<string, OpenFileDialog>();

        public static string OpenAndSave(string name)
        {
            OpenFileDialog ofd = Helper.OpenFile();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (files.ContainsKey(name))
                {
                    files.Remove(name);
                    files.Add(name, ofd);
                }
                else files.Add(name, ofd);
                return ofd.SafeFileName;
            }
            else
            {
                if (files.ContainsKey(name))
                {
                    files.Remove(name);
                }
                return "Error!, Please Try Again";
            }
        }

        public static void Process()
        {
            ExcelProcess.Process(files);
        }

        public static bool CheckFiles()
        {
            //string[] fileNames = { "10Besar", "Kematian", "Hidup", "Laporan" };
            string[] fileNames = { "10Besar", "Hidup", "Laporan" };
            bool isComplete = true;
            foreach(var name in fileNames)
            {
                if (!files.ContainsKey(name))
                {
                    isComplete = false;
                }
            }
            return isComplete;
        }

        public static void EmptyFilesDictionary()
        {
            files = new Dictionary<string, OpenFileDialog>();
        }
    }
}
