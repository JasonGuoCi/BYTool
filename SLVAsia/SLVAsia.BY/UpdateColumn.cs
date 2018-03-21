
using Com.Common.Utility;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLVAsia.BY
{
    public class UpdateColumn
    {
        private static string logPath;
        private static string logPathSuccess;

        private static string importOrderPath;
        private static string importPaymentPath;
        private static string exportPaymentPath;

        public static void UpdateCode()
        {
            logPath = ConfigurationManager.AppSettings["logPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
            logPathSuccess = ConfigurationManager.AppSettings["logPathSuccess"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
            importOrderPath = ConfigurationManager.AppSettings["orderPath"];
            importPaymentPath = ConfigurationManager.AppSettings["paymentPath"];
            exportPaymentPath = ConfigurationManager.AppSettings["exportPaymentPath"];

            Console.WriteLine("Starting..." + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            LogHelper.WriteLogSuccess("Starting..." + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), logPathSuccess);

            //DataTable dt = ReBuildDT();
            DataTable dt = ReBuildDT2();
            LogHelper.WriteLogSuccess("Rebuild Payment datatable success...", logPathSuccess);

            bool flag = ExcelHelper.DataTableToExcel(dt, exportPaymentPath);
            if (flag)
            {
                Console.WriteLine("Rebuild success...");
                LogHelper.WriteLogSuccess("Rebuild success...", logPathSuccess);
            }
            else
            {
                Console.WriteLine("Failure...Please check the error log");
                LogHelper.WriteLog("Failure...", logPath);
            }


        }

        /// <summary>
        /// refresh the dt data
        /// </summary>
        /// <returns></returns>
        public static DataTable ReBuildDT()
        {
            DataTable PaymentDT = ExcelHelper.ExcelToDataTable(importPaymentPath, true, logPath, logPathSuccess);
            Hashtable orderHT = ExcelHelper.ExcelToHashTable(importOrderPath, true, logPath, logPathSuccess);

            //test for exceltohashtable
            LogHelper.WriteLogSuccess("Get datat from Excel success...", logPathSuccess);

            foreach (DataRow row in PaymentDT.Rows)
            {
                string orderNo = @"订单/受理单号";
                string code = row[orderNo].ToString() + "\t";
                if (orderHT.Contains(code))
                {
                    row["商品编码"] = orderHT[code];
                }

            }
            return PaymentDT;
        }

        public static DataTable ReBuildDT2()
        {
            DataTable PaymentDT = ExcelHelper.ExcelToDataTable(importPaymentPath, true, logPath, logPathSuccess);
            DataTable orderDT = ExcelHelper.ExcelToDataTable(importOrderPath, true, logPath, logPathSuccess);

            //test for exceltohashtable
            LogHelper.WriteLogSuccess("Get datat from Excel success...", logPathSuccess);

            foreach (DataRow row in PaymentDT.Rows)
            {
                string orderNo = @"订单/受理单号";
                string code = row[orderNo].ToString() + "\t";
                DataRow[] rows = orderDT.Select("订单号='" + code + "'");
                if (rows.Length > 1)
                {
                    int totalPrice = Convert.ToInt32(row["商品总额"].ToString().Split('.')[0]);
                    int num = Convert.ToInt32(row["商品数量"].ToString().Split('.')[0]);
                    int price = totalPrice / num;
                    string basicCode = string.Empty;
                    for (int i = 0; i < rows.Length; i++)
                    {
                        if (price.ToString() == rows[i][25].ToString().Split('.')[0])
                        {
                            basicCode = rows[i][22].ToString();
                        }
                    }
                    row["商品编码"] = basicCode;

                }
                else if (rows.Length == 1)
                {
                    //LogHelper.WriteLogSuccess(row[0].ToString(), logPathSuccess);
                    row["商品编码"] = rows[0][22].ToString();
                }

            }
            return PaymentDT;
        }
    }
}
