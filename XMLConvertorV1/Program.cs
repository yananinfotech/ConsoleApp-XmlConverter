using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XMLConvertorV1
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            do
            {
                Console.WriteLine("Welcome to XML generator, We support only .xls/.xlsx");
                Console.WriteLine("Press any key and choose a file from explorer");
                Console.ReadKey();
                var fileName = GetFileName();
                if (!string.IsNullOrEmpty(fileName))
                {
                    var dt = ExcelToDatatable.ExcelToTable(fileName);
                    if (dt != null)
                    {
                        //Convert DT to XML
                        DataTableToXml.GenerateXML(dt);
                    }
                    else
                    {

                    }
                }
                else
                    Console.WriteLine("File Not found, Please choose any file");
            } while (true);

        }

        public static string GetFileName()
        {
            string fileName = "";
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileName = ofd.FileName;
                return fileName;
            }
            return "";
        }
    }
}
