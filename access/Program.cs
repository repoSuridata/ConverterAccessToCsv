using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows;

namespace access
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] files = read_dir(ConfigurationManager.AppSettings["originDir"]);
            foreach (string file in files)
            {               
                converter(file);
            }                       
        }

        static void converter(string fileDir)
        {
            try
            {
                OleDbConnection conn = new OleDbConnection();
                DataSet ds = new DataSet();
                OleDbDataAdapter da;
                string newFile = fileDir.Replace("mdb", "csv");
                string bdSource = fileDir;
                string sql = "SELECT * FROM " + ConfigurationManager.AppSettings["bdTableName"];

                string outputCSV = "";

                conn.ConnectionString = ConfigurationManager.AppSettings["connectionString"] + bdSource;
                conn.Open();

                da = new OleDbDataAdapter(sql, conn);
                da.Fill(ds, ConfigurationManager.AppSettings["bdTableName"]);
                int MaxRows = Convert.ToInt32(ds.Tables[0].Rows.Count);
                int NumColumns = Convert.ToInt32(ds.Tables[0].Columns.Count);
                conn.Close();
                int percentage = 1;
                int a = 1;
                Console.WriteLine("Convertendo arquivo:");
                Console.WriteLine(fileDir);
                outputCSV = get_header(ds) + "\n";
                for (int i = 0; i < MaxRows; i++)
                {
                    percentage = (100 * i) / MaxRows;
                    if (percentage == Convert.ToInt32(a + "0"))
                    {
                        Console.WriteLine(percentage + "%");
                        a = a + 1;
                    }
                    for (int j = 0; j < NumColumns; j++)
                    {
                        outputCSV = outputCSV + ds.Tables["ContasPagas"].Rows[i].ItemArray[j] + ";".Trim();
                    }
                    outputCSV = outputCSV + "\r\n";
                }

                create_file(newFile, outputCSV);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        static string get_header(DataSet ds) 
        {
            try
            {
                string header = "";
                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        header = header + column.ColumnName + ";".Trim();
                    }
                }
                return header;
            }
            catch (Exception e) 
            {
                Console.WriteLine(e);
                return null;
            }
        }
        static void create_file(string newFile, string content)
        {
            try
            {
                StreamWriter writer = new StreamWriter(newFile);

                writer.WriteLine(content);

                writer.Close();
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
        }

        static string[] read_dir(string dir) 
        {
            try
            {
                string[] filesDir = Directory.GetFiles(dir);
                return filesDir;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
    }
}