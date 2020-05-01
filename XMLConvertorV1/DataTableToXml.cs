using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XMLConvertorV1
{
    static class DataTableToXml
    {
        public static void GenerateXML(DataTable dataTable)
        {
            //Do XML Conversion here
            Console.WriteLine("Ready to Generate XML \n");
            Console.WriteLine("Please Enter any name for XML \n");

            var xmlFile = Console.ReadLine();
            var xmlFileName = AppDomain.CurrentDomain.BaseDirectory + xmlFile + ".xml";
            try
            {
                //XmlWriterSettings settings = new XmlWriterSettings();
                //settings.Encoding = Encoding.UTF8;
                //using (XmlWriter writer = XmlWriter.Create(xmlFileName, settings))
                //{
                //    dataTable.WriteXml(writer);
                //    writer.Close();
                //}

                using (StreamWriter fs = new StreamWriter(xmlFileName)) // XML File Path
                {
                    dataTable.WriteXml(fs);
                }
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Hurray!!!!!!!!!!! XML Generated, \n Your XML file is here:\n {xmlFileName}");

                Console.ResetColor();
                Console.WriteLine("\n \n Want to generate another XML, Press any key!");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Message: {ex.Message} , Stacktrace: {ex.StackTrace}");
                Console.WriteLine("Something Went in writing XML, Please try with some other file");
            }
        }
    }
}
