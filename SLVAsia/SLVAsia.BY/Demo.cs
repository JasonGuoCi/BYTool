using Com.Common.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLVSPDataExport
{
    class Program
    {
        private static string logPath;
        private static string logPathSuccess;
        private static string sourceSPListFields;
        private static string resultPath;
        private static string headerName = string.Empty;
        //static void Main(string[] args)
        //{
        //    Console.WriteLine("Export Data Start...");
        //    logPath = ConfigurationManager.AppSettings["logPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
        //    logPathSuccess = ConfigurationManager.AppSettings["logPathSuccess"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
        //    sourceSPListFields = ConfigurationManager.AppSettings["sourceSPListFields"];
        //}

        public static DataTable ReBuildDT()
        {
            DataTable sourceSPListFieldsDT = ExcelHelper.ExcelToDataTable(sourceSPListFields, true, logPath, logPathSuccess);

            if (sourceSPListFieldsDT != null && sourceSPListFieldsDT.Rows.Count > 0)
            {
                //把header在csv中生成
                for (int i = 1; i < sourceSPListFieldsDT.Rows.Count; i++)
                {
                    headerName += sourceSPListFieldsDT.Rows[i]["Header"].ToString() + ",";
                    //for (int j = 0; j < sourceSPListFieldsDT.Columns.Count; j++)
                    //{
                    //    headerName += sourceSPListFieldsDT.Rows[i]["Header"].ToString() + ",";
                    //}
                }
                //WriteResult(headerName);

            }
            return null;
        }

        //public static void WriteResult(string text)
        //{
        //    FileStream fs = new FileStream(resultPath, FileMode.Append);
        //    StreamWriter sw = new StreamWriter(fs, Encoding.Default);
        //    sw.WriteLine(text);
        //    sw.Close();
        //    fs.Close();
        //}
    }
}
