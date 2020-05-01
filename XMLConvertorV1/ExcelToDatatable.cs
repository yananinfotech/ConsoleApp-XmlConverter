using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLConvertorV1
{
    static class ExcelToDatatable
    {
        public static DataTable ExcelToTable(string fileName)
        {
            Console.WriteLine("Reading your Excel Content");

            try
            {
                DataTable dt = new DataTable();
                string Ext = Path.GetExtension(fileName);
                string connectionString = "";

                if (Ext == ".xls")
                {   //For Excel 97-03  
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source =" + fileName + "; Extended Properties = 'Excel 8.0;HDR=YES'";
                }
                else if (Ext == ".xlsx")
                {    //For Excel 07 and greater  
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source =" + fileName + "; Extended Properties = 'Excel 8.0;HDR=YES'";
                }
                else
                {
                    Console.BackgroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("File format not supported");
                    Console.ResetColor();
                    return null;
                }

                OleDbConnection conn = new OleDbConnection(connectionString);
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter();

                cmd.Connection = conn;
                //Fetch Fisrt Sheet Name  
                conn.Open();
                DataTable dtSchema;
                dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string ExcelSheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
                dtSchema.TableName = ExcelSheetName;
                conn.Close();
                //Read all data from the Sheet to a Data Table  
                conn.Open();
                cmd.CommandText = "SELECT * From [" + ExcelSheetName + "]";
                dataAdapter.SelectCommand = cmd;
                dataAdapter.Fill(dt); // Fill Sheet Data to Datatable  
                conn.Close();
                dt.TableName = ExcelSheetName.Replace("$", "");
                Console.WriteLine("Successfully read your content in Excel!");
                return dt;
            }
            catch
            {
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Cannot Read Content, Please try with other file.");
                Console.ResetColor();
                return null;
            }
        }
    }
}
