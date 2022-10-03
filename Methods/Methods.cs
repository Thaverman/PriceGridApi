using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Runtime.InteropServices;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using PriceGridApi.Models.UCommerceModels;
using PriceGridApi.Models;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;

namespace PriceGridApi.Methods
{
    public class Methods
    {        
        readonly string CSPG = "Server=SSWTHINKPAD51\\MSSQLSERVER01;User Id=; Password=;Database=SSW_cs82PriceGrid"; // Uded with shipping data from PriceGrid 
        readonly string CS = "Server=VMAPP05;User Id=NxtAnalyst2; Password=NxtAnalyst2;Database=nxt_views";
        private void WriteToFile(string text)
        {
            string path = @"C:\Users\thaverman\Documents\DuplicateOrderSearch.txt";
            using (System.IO.StreamWriter writer = File.AppendText(path))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.WriteLine();
                writer.Close();
            }
        }
        public void GetOrderLines(Dictionary<string, UCommerce_OrderLine> UCommerce_OrderLines, string sql) //Fills a dictionary with UCommerce_OrderLines Line Items information using OrderLineId as the Key
        {
            SqlConnection cont = null;
            SqlDataReader rdrR = null;

            try
            {
                using (cont = new SqlConnection(CS))
                {
                    cont.Open();
                    SqlCommand cmd1 = new SqlCommand(sql, cont);

                    cmd1.CommandType = CommandType.Text;

                    rdrR = cmd1.ExecuteReader();

                    while (rdrR.Read())
                    {
                        UCommerce_OrderLine uCommerce_OrderLine = new UCommerce_OrderLine();

                        var outputParam1 = rdrR["OrderLineId"];
                        if (!(outputParam1 is DBNull))
                        {
                            uCommerce_OrderLine.OrderLineId = Convert.ToInt32(outputParam1);
                        }
                        else
                        {
                            uCommerce_OrderLine.OrderLineId = -96;
                        }

                        var outputParam2 = rdrR["OrderId"];
                        if (!(outputParam2 is DBNull))
                        {
                            uCommerce_OrderLine.OrderId = Convert.ToInt32(outputParam2);
                        }
                        else
                        {
                            uCommerce_OrderLine.OrderId = -05810;
                        }

                        var outputParam3 = rdrR["Sku"];
                        if (!(outputParam3 is DBNull))
                        {
                            uCommerce_OrderLine.Sku = Convert.ToString(outputParam3);
                        }
                        else
                        {
                            uCommerce_OrderLine.Sku = null;
                        }

                        var outputParam4 = rdrR["ProductName"];
                        if (!(outputParam4 is DBNull))
                        {
                            uCommerce_OrderLine.ProductName = Convert.ToString(outputParam4);
                        }
                        else
                        {
                            uCommerce_OrderLine.ProductName = null;
                        }

                        var outputParam5 = rdrR["Price"];
                        if (!(outputParam5 is DBNull))
                        {
                            uCommerce_OrderLine.Price = Convert.ToDecimal(outputParam5);
                        }
                        else
                        {
                            uCommerce_OrderLine.Price = null;
                        }

                        var outputParam6 = rdrR["Quantity"];
                        if (!(outputParam6 is DBNull))
                        {
                            uCommerce_OrderLine.Quantity = Convert.ToInt32(outputParam6);
                        }
                        else
                        {
                            uCommerce_OrderLine.Quantity = -96;
                        }

                        var outputParam7 = rdrR["CreatedOn"];
                        if (!(outputParam7 is DBNull))
                        {
                            uCommerce_OrderLine.CreatedOn = Convert.ToDateTime(outputParam7);
                        }
                        else
                        {
                            uCommerce_OrderLine.CreatedOn = null;
                        }

                        var outputParam8 = rdrR["Discount"];
                        if (!(outputParam8 is DBNull))
                        {
                            uCommerce_OrderLine.Discount = Convert.ToDecimal(outputParam8);
                        }
                        else
                        {
                            uCommerce_OrderLine.Discount = null;
                        }

                        var outputParam9 = rdrR["VAT"];
                        if (!(outputParam9 is DBNull))
                        {
                            uCommerce_OrderLine.VAT = Convert.ToDecimal(outputParam9);
                        }
                        else
                        {
                            uCommerce_OrderLine.VAT = null;
                        }

                        var outputParam10 = rdrR["Total"];
                        if (!(outputParam10 is DBNull))
                        {
                            uCommerce_OrderLine.Total = Convert.ToDecimal(outputParam10);
                        }
                        else
                        {
                            uCommerce_OrderLine.Total = null;
                        }

                        var outputParam11 = rdrR["VATRate"];
                        if (!(outputParam11 is DBNull))
                        {
                            uCommerce_OrderLine.VATRate = Convert.ToDecimal(outputParam11);
                        }
                        else
                        {
                            uCommerce_OrderLine.VATRate = null;
                        }

                        var outputParam12 = rdrR["VariantSku"];
                        if (!(outputParam12 is DBNull))
                        {
                            uCommerce_OrderLine.VariantSku = Convert.ToString(outputParam12);
                        }
                        else
                        {
                            uCommerce_OrderLine.VariantSku = null;
                        }

                        var outputParam13 = rdrR["ShipmentId"];
                        if (!(outputParam13 is DBNull))
                        {
                            uCommerce_OrderLine.ShipmentId = Convert.ToInt32(outputParam13);
                        }
                        else
                        {
                            uCommerce_OrderLine.ShipmentId = null;
                        }

                        var outputParam14 = rdrR["UnitDiscount"];
                        if (!(outputParam14 is DBNull))
                        {
                            uCommerce_OrderLine.UnitDiscount = Convert.ToDecimal(outputParam14);
                        }
                        else
                        {
                            uCommerce_OrderLine.UnitDiscount = null;
                        }

                        var outputParam15 = rdrR["CreatedBy"];
                        if (!(outputParam15 is DBNull))
                        {
                            uCommerce_OrderLine.CreatedBy = Convert.ToString(outputParam15);
                        }
                        else
                        {
                            uCommerce_OrderLine.CreatedBy = null;
                        }

                        if (uCommerce_OrderLine.Total.HasValue && uCommerce_OrderLine.Total != null && uCommerce_OrderLine.Total > 0)
                        {
                            UCommerce_OrderLines.Add(uCommerce_OrderLine.OrderId.ToString() + uCommerce_OrderLine.OrderLineId + uCommerce_OrderLine.Sku, uCommerce_OrderLine);
                        }

                    }
                    cont.Close();
                    rdrR.Close();
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
            }
            finally
            {
                if (cont != null)
                {
                    cont.Close();
                }
                if (rdrR != null)
                {
                    rdrR.Close();
                }
            }


            //}

        }
        //public void GetExcelFile(string[,] excelArray) // This was based on an excel report that Price Grid Generated with Shipping infomation. The shipping information was very sparce
        //{


        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\thaverman\Documents\200 Products 10.5.21.xlsx");
        //    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Excel.Range xlRange = xlWorksheet.UsedRange;

        //    int rowCount = xlRange.Rows.Count;
        //    int colCount = xlRange.Columns.Count;


        //    for (int i = 1; i <= rowCount; i++)
        //    {
        //        for (int j = 1; j <= colCount; j++)
        //        {
        //            if (j - 1 < 2)
        //            {
        //                excelArray[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();
        //            }


        //            //if (j == 1)
        //            //    Console.Write("\r\n");

        //            //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
        //            //{
        //            //    Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

        //            //}



        //        }
        //    }


        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();


        //    Marshal.ReleaseComObject(xlRange);
        //    Marshal.ReleaseComObject(xlWorksheet);


        //    xlWorkbook.Close();
        //    Marshal.ReleaseComObject(xlWorkbook);


        //    xlApp.Quit();
        //    Marshal.ReleaseComObject(xlApp);
        //}
        //public void GetPriceGrid(Dictionary<string, Sheet> Checks, string sql, int typeOfKey)
        //{
        //    SqlConnection cont = null;
        //    SqlDataReader rdrR = null;




        //    try
        //    {
        //        using (cont = new SqlConnection(CSPG))
        //        {
        //            cont.Open();
        //            SqlCommand cmd1 = new SqlCommand(sql, cont);




        //            cmd1.CommandType = CommandType.Text;

        //            rdrR = cmd1.ExecuteReader();

        //            while (rdrR.Read())
        //            {
        //                Sheet Check = new Sheet();


        //                var outputParam1 = rdrR["id"];
        //                if (!(outputParam1 is DBNull))
        //                {
        //                    Check.Id = Convert.ToInt32(outputParam1);
        //                }
        //                else
        //                {
        //                    Check.Id = -96;
        //                }

        //                var outputParam99 = rdrR["Brand"];
        //                if (!(outputParam99 is DBNull))
        //                {
        //                    Check.Brand = Convert.ToString(outputParam99);
        //                }
        //                else
        //                {
        //                    Check.Brand = null;
        //                }


        //                var outputParam2 = rdrR["Brand SKU"];
        //                if (!(outputParam2 is DBNull))
        //                {
        //                    Check.Brand_SKU = Convert.ToString(outputParam2);
        //                }
        //                else
        //                {
        //                    Check.Brand_SKU = null;
        //                }

        //                var outputParam3 = rdrR["Product Name"];
        //                if (!(outputParam3 is DBNull))
        //                {
        //                    Check.Product_Name = Convert.ToString(outputParam3);
        //                }
        //                else
        //                {
        //                    Check.Product_Name = null;
        //                }

        //                var outputParam4 = rdrR["Our Price"];
        //                if (!(outputParam4 is DBNull))
        //                {
        //                    Check.Our_Price = Convert.ToDecimal(outputParam4);
        //                }
        //                else
        //                {
        //                    Check.Our_Price = -96;
        //                }

        //                var outputParam5 = rdrR["Product ID"];
        //                if (!(outputParam5 is DBNull))
        //                {
        //                    Check.Product_Id = Convert.ToString(outputParam5);
        //                }
        //                else
        //                {
        //                    Check.Product_Id = null;
        //                }

        //                var outputParam6 = rdrR["Average Store Price"];
        //                if (!(outputParam6 is DBNull))
        //                {
        //                    Check.Average_Store_Price = Convert.ToDecimal(outputParam6);
        //                }
        //                else
        //                {
        //                    Check.Average_Store_Price = -96;
        //                }

        //                var outputParam7 = rdrR["Max Store Price"];
        //                if (!(outputParam7 is DBNull))
        //                {
        //                    Check.Max_Store_Price = Convert.ToDecimal(outputParam7);
        //                }
        //                else
        //                {
        //                    Check.Max_Store_Price = null;
        //                }

        //                var outputParam8 = rdrR["Median Store Price"];
        //                if (!(outputParam8 is DBNull))
        //                {
        //                    Check.Median_Store_Price = Convert.ToDecimal(outputParam8);
        //                }
        //                else
        //                {
        //                    Check.Median_Store_Price = null;
        //                }

        //                var outputParam9 = rdrR["Min Store Price"];
        //                if (!(outputParam9 is DBNull))
        //                {
        //                    Check.Min_Store_Price = Convert.ToDecimal(outputParam9);
        //                }
        //                else
        //                {
        //                    Check.Min_Store_Price = null;
        //                }

        //                var outputParam10 = rdrR["Store Name"];
        //                if (!(outputParam10 is DBNull))
        //                {
        //                    Check.Store_Name = Convert.ToString(outputParam10);
        //                }
        //                else
        //                {
        //                    Check.Store_Name = null;
        //                }

        //                var outputParam11 = rdrR["Store Price"];
        //                if (!(outputParam11 is DBNull))
        //                {
        //                    Check.Store_Price = Convert.ToDecimal(outputParam11);
        //                }
        //                else
        //                {
        //                    Check.Store_Price = 0m;
        //                }

        //                var outputParam12 = rdrR["Store Shipping Price (01331)"];
        //                if (!(outputParam12 is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_01331 = Convert.ToString(outputParam12);
        //                }
        //                else
        //                {
        //                    Check.Store_Shipping_Price_01331 = null;
        //                }

        //                var outputParam13 = rdrR["Store Shipping Price (01364)"];
        //                if (!(outputParam13 is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_01364 = Convert.ToString(outputParam13);
        //                }
        //                else
        //                {
        //                    Check.Store_Shipping_Price_01364 = null;
        //                }

        //                var outputParam14 = rdrR["Store Shipping Price (01378)"];
        //                if (!(outputParam14 is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_01378 = Convert.ToString(outputParam14);
        //                }
        //                else
        //                {
        //                    Check.Store_Shipping_Price_01378 = null;
        //                }

        //                var outputParamDecmial = rdrR["Store Shipping Price (06460)"];
        //                if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_06460 = Convert.ToString(outputParamDecmial);
        //                }
        //                else
        //                {
        //                    Check.Store_Shipping_Price_06460 = null;
        //                }

        //                outputParamDecmial = rdrR["Store Shipping Price (06610)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_06610 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_06610 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (06615)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_06615 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_06615 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (30297)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_30297 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_30297 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (30304)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_30304 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_30304 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (30354)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_30354 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_30354 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (34114)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_34114 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_34114 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (34116)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_34116 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_34116 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (34117)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_34117 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_34117 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (37208)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_37208 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_37208 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (37209)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_37209 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_37209 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (37217)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_37217 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_37217 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (45312)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_45312 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_45312 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (45371)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_45371 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_45371 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (45373)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_45373 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_45373 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (46809)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_46809 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_46809 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (46818)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_46818 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_46818 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (46855)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_46855 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_46855 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (48192)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_48192 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_48192 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (48193)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_48193 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_48193 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (48195)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_48195 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_48195 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (50311)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_50311 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_50311 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (50312)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_50312 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_50312 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (50313)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_50313 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_50313 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (55401)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_55401 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_55401 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (55403)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_55403 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_55403 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (55421)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_55421 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_55421 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (59601)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_59601 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_59601 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (59602)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_59602 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_59602 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (59635)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_59635 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_59635 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (60602)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_60602 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_60602 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (60608)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_60608 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_60608 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (60661)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_60661 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_60661 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (63044)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_63044 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_63044 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (67207)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_67207 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_67207 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (67208)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_67208 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_67208 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (67214)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_67214 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_67214 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (67228)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_67228 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_67228 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (68136)"]; if (!(outputParamDecmial is DBNull))
        //                {
        //                    Check.Store_Shipping_Price_68136 = Convert.ToString(outputParamDecmial);
        //                }
        //                else { Check.Store_Shipping_Price_68136 = null; }

        //                outputParamDecmial = rdrR["Store Shipping Price (68137)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_68137 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_68137 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (68138)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_68138 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_68138 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (73110)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_73110 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_73110 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (73115)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_73115 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_73115 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (73129)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_73129 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_73129 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (77011)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_77011 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_77011 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (77023)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_77023 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_77023 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (77044)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_77044 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_77044 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (78502)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_78502 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_78502 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (78504)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_78504 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_78504 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (78505)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_78505 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_78505 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (80229)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_80229 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_80229 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (80241)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_80241 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_80241 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (80260)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_80260 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_80260 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (82923)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_82923 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_82923 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (82930)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_82930 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_82930 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (82931)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_82931 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_82931 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (82932)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_82932 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_82932 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (82941)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_82941 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_82941 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (83706)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_83706 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_83706 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (83714)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_83714 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_83714 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (83799)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_83799 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_83799 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (84741)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_84741 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_84741 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (84755)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_84755 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_84755 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (84758)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_84758 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_84758 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85002)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85002 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85002 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85003)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85003 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85003 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85004)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85004 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85004 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85713)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85713 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85713 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85719)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85719 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85719 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (85726)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_85726 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_85726 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (87501)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_87501 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_87501 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (87505)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_87505 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_87505 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (87507)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_87507 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_87507 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (89012)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_89012 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_89012 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (89015)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_89015 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_89015 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (89030)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_89030 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_89030 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (89115)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_89115 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_89115 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (89180)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_89180 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_89180 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (90501)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_90501 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_90501 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (90505)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_90505 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_90505 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (90710)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_90710 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_90710 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (92802)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_92802 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_92802 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (92861)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_92861 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_92861 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (92865)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_92865 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_92865 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (94115)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_94115 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_94115 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (94118)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_94118 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_94118 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (94122)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_94122 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_94122 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (98108)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_98108 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_98108 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (98134)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_98134 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_98134 = null; }
        //                outputParamDecmial = rdrR["Store Shipping Price (98144)"]; if (!(outputParamDecmial is DBNull)) { Check.Store_Shipping_Price_98144 = Convert.ToString(outputParamDecmial); } else { Check.Store_Shipping_Price_98144 = null; }

        //                var outputParamDate = rdrR["Store Price Date"]; if (!(outputParamDate is DBNull)) { Check.Store_Price_Date = Convert.ToDateTime(outputParamDate); } else { Check.Store_Price_Date = null; }


        //                if (Check.Store_Price != 0m && typeOfKey == 0)
        //                {
        //                    Checks.Add(Check.Store_Price + Check.Id.ToString() + Check.Brand_SKU, Check);
        //                }
        //                else if (Check.Store_Price != 0m && typeOfKey == 1)
        //                {
        //                    Checks.Add(Check.Brand_SKU, Check);
        //                }
        //                else if (Check.Store_Price != 0m && typeOfKey == 2)
        //                {
        //                    Checks.Add(Check.Brand_SKU + Check.Store_Name, Check);
        //                }








        //            }
        //            cont.Close();
        //            rdrR.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
        //    }
        //    finally
        //    {
        //        if (cont != null)
        //        {
        //            cont.Close();
        //        }
        //        if (rdrR != null)
        //        {
        //            rdrR.Close();
        //        }
        //    }


        //    //}

        //}
        public void GetDuplicatePurchaseOrders(Dictionary<string, UCommerce_PurchaseOrder> checker, string sql)//Fills a dictionary with PurchaseOrder information using orderId as the Key
        {
            SqlConnection cont = null;
            SqlDataReader rdrR = null;


            DateTime NoData = new DateTime(1753, 01, 01, 12, 00, 00);







            try
            {
                using (cont = new SqlConnection(CS))
                {
                    cont.Open();
                    SqlCommand cmd1 = new SqlCommand(sql, cont);




                    cmd1.CommandType = CommandType.Text;

                    rdrR = cmd1.ExecuteReader();

                    while (rdrR.Read())
                    {

                        UCommerce_PurchaseOrder UCommerce_PurchaseOrder = new UCommerce_PurchaseOrder();

                        var outputParam1 = rdrR["OrderId"];
                        if (!(outputParam1 is DBNull))
                        {
                            UCommerce_PurchaseOrder.OrderId = Convert.ToInt32(outputParam1);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.OrderId = null;
                        }


                        var outputParam2 = rdrR["OrderNumber"];
                        if (!(outputParam2 is DBNull))
                        {
                            UCommerce_PurchaseOrder.OrderNumber = Convert.ToString(outputParam2);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.OrderNumber = null;
                        }

                        var outputParam3 = rdrR["CustomerId"];
                        if (!(outputParam3 is DBNull))
                        {
                            UCommerce_PurchaseOrder.CustomerId = Convert.ToInt32(outputParam3);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.CustomerId = null;
                        }

                        var outputParam4 = rdrR["OrderStatusId"];
                        if (!(outputParam4 is DBNull))
                        {
                            UCommerce_PurchaseOrder.OrderStatusId = Convert.ToInt32(outputParam4);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.OrderStatusId = null;
                        }

                        var outputParam5 = rdrR["CreatedDate"];
                        if (!(outputParam5 is DBNull))
                        {
                            UCommerce_PurchaseOrder.CreatedDate = Convert.ToDateTime(outputParam5);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.CreatedDate = NoData;
                        }

                        var outputParam6 = rdrR["CompletedDate"];
                        if (!(outputParam6 is DBNull))
                        {
                            UCommerce_PurchaseOrder.CompletedDate = Convert.ToDateTime(outputParam6);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.CompletedDate = null;
                        }

                        var outputParam7 = rdrR["CurrencyId"];
                        if (!(outputParam7 is DBNull))
                        {
                            UCommerce_PurchaseOrder.CurrencyId = Convert.ToInt32(outputParam7);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.CurrencyId = null;
                        }


                        var outputParam8 = rdrR["ProductCatalogGroupId"];
                        if (!(outputParam8 is DBNull))
                        {
                            UCommerce_PurchaseOrder.ProductCatalogGroupId = Convert.ToInt32(outputParam8);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.ProductCatalogGroupId = null;
                        }

                        var outputParam9 = rdrR["BillingAddressId"];
                        if (!(outputParam9 is DBNull))
                        {
                            UCommerce_PurchaseOrder.BillingAddressId = Convert.ToInt32(outputParam9);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.BillingAddressId = null;
                        }

                        var outputParam10 = rdrR["Note"];
                        if (!(outputParam10 is DBNull))
                        {
                            UCommerce_PurchaseOrder.Note = Convert.ToString(outputParam10);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.Note = null;
                        }

                        var outputParam11 = rdrR["BasketId"];
                        if (!(outputParam11 is DBNull))
                        {
                            UCommerce_PurchaseOrder.BasketId = Convert.ToString(outputParam11);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.BasketId = null;
                        }

                        var outputParam12 = rdrR["VAT"];
                        if (!(outputParam12 is DBNull))
                        {
                            UCommerce_PurchaseOrder.VAT = Convert.ToInt32(outputParam12);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.VAT = null;
                        }

                        var outputParam13 = rdrR["OrderTotal"];
                        if (!(outputParam13 is DBNull))
                        {
                            UCommerce_PurchaseOrder.OrderTotal = Convert.ToDecimal(outputParam13);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.OrderTotal = null;
                        }

                        var outputParam14 = rdrR["ShippingTotal"];
                        if (!(outputParam14 is DBNull))
                        {
                            UCommerce_PurchaseOrder.ShippingTotal = Convert.ToDecimal(outputParam14);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.ShippingTotal = null;
                        }

                        var outputParam15 = rdrR["PaymentTotal"];
                        if (!(outputParam15 is DBNull))
                        {
                            UCommerce_PurchaseOrder.PaymentTotal = Convert.ToDecimal(outputParam15);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.PaymentTotal = null;
                        }

                        var outputParam16 = rdrR["TaxTotal"];
                        if (!(outputParam16 is DBNull))
                        {
                            UCommerce_PurchaseOrder.TaxTotal = Convert.ToInt32(outputParam16);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.TaxTotal = null;
                        }

                        var outputParam17 = rdrR["SubTotal"];
                        if (!(outputParam17 is DBNull))
                        {
                            UCommerce_PurchaseOrder.SubTotal = Convert.ToDecimal(outputParam17);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.SubTotal = null;
                        }

                        var outputParam18 = rdrR["OrderGuid"];
                        if (!(outputParam18 is DBNull))
                        {
                            UCommerce_PurchaseOrder.OrderGuid = Convert.ToString(outputParam18);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.OrderGuid = null;
                        }

                        var outputParam19 = rdrR["ModifiedOn"];
                        if (!(outputParam19 is DBNull))
                        {
                            UCommerce_PurchaseOrder.ModifiedOn = Convert.ToDateTime(outputParam19);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.ModifiedOn = null;
                        }

                        var outputParam20 = rdrR["CultureCode"];
                        if (!(outputParam20 is DBNull))
                        {
                            UCommerce_PurchaseOrder.CultureCode = Convert.ToString(outputParam20);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.CultureCode = null;
                        }

                        var outputParam21 = rdrR["Discount"];
                        if (!(outputParam21 is DBNull))
                        {
                            UCommerce_PurchaseOrder.Discount = Convert.ToDecimal(outputParam21);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.Discount = null;
                        }

                        var outputParam22 = rdrR["DiscountTotal"];
                        if (!(outputParam22 is DBNull))
                        {
                            UCommerce_PurchaseOrder.DiscountTotal = Convert.ToInt32(outputParam22);
                        }
                        else
                        {
                            UCommerce_PurchaseOrder.DiscountTotal = null;
                        }


                        if (UCommerce_PurchaseOrder.CompletedDate != null && UCommerce_PurchaseOrder.OrderTotal != null) //&& string.UCommerce_PurchaseOrder.Note.("Test")
                        {
                            if (UCommerce_PurchaseOrder.SubTotal > 0.0m)
                            {
                                if (string.IsNullOrEmpty(UCommerce_PurchaseOrder.Note))
                                {
                                    checker.Add(UCommerce_PurchaseOrder.OrderId.ToString(), UCommerce_PurchaseOrder);
                                }
                                else if (!UCommerce_PurchaseOrder.Note.ToLower().Contains("test order"))
                                {
                                    checker.Add(UCommerce_PurchaseOrder.OrderId.ToString(), UCommerce_PurchaseOrder);
                                }
                            }


                        }



                    }
                    cont.Close();
                    rdrR.Close();
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
            }
            finally
            {
                if (cont != null)
                {
                    cont.Close();
                }
                if (rdrR != null)
                {
                    rdrR.Close();
                }
            }




        }
        //public void GetConversionTable(Dictionary<string, Conversions> checker, string sql)//I was maintaining a conversion table in a local sql DB the information is very out of date
        //{
        //    SqlConnection cont = null;
        //    SqlDataReader rdrR = null;


        //    try
        //    {
        //        using (cont = new SqlConnection(CSPG))
        //        {
        //            cont.Open();
        //            SqlCommand cmd1 = new SqlCommand(sql, cont);




        //            cmd1.CommandType = CommandType.Text;

        //            rdrR = cmd1.ExecuteReader();

        //            while (rdrR.Read())
        //            {

        //                Conversions tempAdd = new Conversions();

        //                var outputParamString = rdrR["SkuCompetitorNameID"];
        //                if (!(outputParamString is DBNull))
        //                {
        //                    tempAdd.SkuCompetitorNameID = Convert.ToString(outputParamString);
        //                }
        //                else
        //                {
        //                    tempAdd.SkuCompetitorNameID = null;
        //                }

        //                outputParamString = rdrR["Product SKU"];
        //                if (!(outputParamString is DBNull))
        //                {
        //                    tempAdd.ProductSKU = Convert.ToString(outputParamString);
        //                }
        //                else
        //                {
        //                    tempAdd.ProductSKU = null;
        //                }

        //                var outputParamDecimal = rdrR["COMP QTY"];
        //                if (!(outputParamDecimal is DBNull))
        //                {
        //                    tempAdd.COMP_QTY = Convert.ToDecimal(outputParamDecimal);
        //                }
        //                else
        //                {
        //                    tempAdd.COMP_QTY = 0.0m;
        //                }

        //                outputParamDecimal = rdrR["SSW QTY"];
        //                if (!(outputParamDecimal is DBNull))
        //                {
        //                    tempAdd.SSW_QTY = Convert.ToDecimal(outputParamDecimal);
        //                }
        //                else
        //                {
        //                    tempAdd.SSW_QTY = 0.0m;
        //                }

        //                outputParamDecimal = rdrR["Conversion"];
        //                if (!(outputParamDecimal is DBNull))
        //                {
        //                    tempAdd.Conversion = Convert.ToDecimal(outputParamDecimal);
        //                }
        //                else
        //                {
        //                    tempAdd.Conversion = 0.0m;
        //                }

        //                outputParamString = rdrR["Competitor Name"];
        //                if (!(outputParamString is DBNull))
        //                {
        //                    tempAdd.Competitor_Name = Convert.ToString(outputParamString);
        //                }
        //                else
        //                {
        //                    tempAdd.Competitor_Name = null;
        //                }

        //                outputParamString = rdrR["Company"];
        //                if (!(outputParamString is DBNull))
        //                {
        //                    tempAdd.Company = Convert.ToString(outputParamString);
        //                }
        //                else
        //                {
        //                    tempAdd.Company = null;
        //                }

        //                var adder = checker.ContainsKey(tempAdd.SkuCompetitorNameID);
        //                if (!adder)
        //                {
        //                    //string temp = tempAdd.SkuCompetitorNameID.ToLower();
        //                    checker.Add(tempAdd.SkuCompetitorNameID, tempAdd);
        //                }

        //            }
        //            cont.Close();
        //            rdrR.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
        //    }
        //    finally
        //    {
        //        if (cont != null)
        //        {
        //            cont.Close();
        //        }
        //        if (rdrR != null)
        //        {
        //            rdrR.Close();
        //        }
        //    }




        //}
        //public void WriteToExcel(string dateStart, Dictionary<string, Report> Reports, int countSheet, int totalCustomers, int totalCustomersGettingRefund, string endDate)
        //{
        //    //This was used to make a small report
        //    string path = "d:\\PriceMatchStart.xlsx";
        //    Application ExcelApp = new Application();
        //    decimal sumRefund = 0.0m; int countRefund = 0;
        //    Workbook ExcelWorkBook = null;

        //    Worksheet ExcelWorkSheet = null;
        //    ExcelApp.Visible = true;

        //    if (countSheet > 1)
        //    {
        //        ExcelWorkBook = ExcelApp.Workbooks.Open(path);

        //        //ExcelWorkBook = ExcelApp.Workbooks.Add();
        //        ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook
        //        //ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //    }
        //    else
        //    {
        //        ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

        //    }

        //    int row = Reports.Count();

        //    try
        //    {
        //        int r = 1; int c = 1;

        //        ExcelWorkSheet = ExcelWorkBook.Worksheets[1];
        //        ExcelWorkSheet.Cells[r, c] = "SKU";
        //        ExcelWorkSheet.Cells[r, c + 1] = "Amount of refund";
        //        ExcelWorkSheet.Cells[r, c + 2] = "Number of refunds";
        //        ExcelWorkSheet.Cells[r, c + 3] = "Number of items";
        //        ExcelWorkSheet.Cells[r, c + 4] = "Customer Price";
        //        ExcelWorkSheet.Cells[r, c + 5] = "Company match name";
        //        ExcelWorkSheet.Cells[r, c + 6] = "Company match price";
        //        ExcelWorkSheet.Cells[r, c + 7] = "Conversion";
        //        ExcelWorkSheet.Cells[r, c + 8] = "Difference";
        //        ExcelWorkSheet.Cells[r, c + 9] = dateStart;
        //        ExcelWorkSheet.Cells[r + 1, c + 9] = endDate;
        //        ExcelWorkSheet.Cells[r, c + 10] = "Total Orders in time Frame";
        //        ExcelWorkSheet.Cells[r, c + 11] = "Total Orders Refunded";
        //        ExcelWorkSheet.Cells[r + 1, c + 10] = totalCustomers;
        //        ExcelWorkSheet.Cells[r + 1, c + 11] = totalCustomersGettingRefund;


        //        r++;

        //        foreach (var item in Reports)
        //        {

        //            ExcelWorkSheet.Cells[r, c] = item.Key;

        //            ExcelWorkSheet.Cells[r, c + 1] = item.Value.TotalRefund;


        //            ExcelWorkSheet.Cells[r, c + 2] = item.Value.count;


        //            ExcelWorkSheet.Cells[r, c + 3] = item.Value.Quantity;


        //            ExcelWorkSheet.Cells[r, c + 4] = item.Value.OurPrice;


        //            ExcelWorkSheet.Cells[r, c + 5] = item.Value.CompName;


        //            ExcelWorkSheet.Cells[r, c + 6] = item.Value.CompPrice;


        //            ExcelWorkSheet.Cells[r, c + 7] = item.Value.Conversion;


        //            ExcelWorkSheet.Cells[r, c + 8] = item.Value.PriceDifference;


        //            sumRefund += item.Value.TotalRefund;
        //            countRefund += item.Value.count;
        //            r++;

        //        }
        //        ExcelWorkSheet.Cells[r, c + 1] = sumRefund;
        //        ExcelWorkSheet.Cells[r, c + 2] = countRefund;

        //        ExcelWorkBook.Worksheets[1].Name = dateStart;//Renaming the Sheet1 to MySheet
        //        if (countSheet < 2)
        //        {
        //            ExcelWorkBook.SaveAs("d:\\PriceMatchStart.xlsx");
        //        }
        //        else
        //        {
        //            ExcelWorkBook.Save();
        //        }
        //        ExcelWorkBook.Close();
        //        ExcelApp.Quit();
        //        Marshal.ReleaseComObject(ExcelWorkSheet);
        //        Marshal.ReleaseComObject(ExcelWorkBook);
        //        Marshal.ReleaseComObject(ExcelApp);
        //    }


        //    //try

        //    //{

        //    //    ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data

        //    //    //Writing data into excel of 100 rows with 10 column 

        //    //    for (int r = 1; r < row; r++) //r stands for ExcelRow and c for ExcelColumn

        //    //    {

        //    //        // Excel row and column start positions for writing Row=1 and Col=1

        //    //        for (int c = 1; c < 9; c++)

        //    //            ExcelWorkSheet.Cells[r, c] = "R" + r + "C" + c;

        //    //    }

        //    //    ExcelWorkBook.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet

        //    //    ExcelWorkBook.SaveAs("d:\\PriceMatch" + dateStart + ".xlsx");

        //    //    ExcelWorkBook.Close();

        //    //    ExcelApp.Quit();

        //    //    Marshal.ReleaseComObject(ExcelWorkSheet);

        //    //    Marshal.ReleaseComObject(ExcelWorkBook);

        //    //    Marshal.ReleaseComObject(ExcelApp);

        //    //}

        //    catch (Exception ex)

        //    {

        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace + ex.GetType());

        //    }

        //    finally

        //    {



        //        foreach (Process process in Process.GetProcessesByName("Excel"))

        //            process.Kill();

        //    }

        //}
        //public void GetMatchSKUPerStore(Dictionary<string, MatchSKUPerStore> checker, string sql)//Used when matching our top 200 items to competitors items and price This used a local DB and not the API
        //{
        //    SqlConnection cont = null;
        //    SqlDataReader rdrR = null;




        //    try
        //    {
        //        using (cont = new SqlConnection(CSPG))
        //        {
        //            cont.Open();
        //            SqlCommand cmd1 = new SqlCommand(sql, cont);




        //            cmd1.CommandType = CommandType.Text;

        //            rdrR = cmd1.ExecuteReader();

        //            while (rdrR.Read())
        //            {
        //                MatchSKUPerStore Match_SKU_Per_Store = new MatchSKUPerStore();

        //                Match_SKU_Per_Store.BrandSKU = (string)rdrR["BrandSKU"];
        //                Match_SKU_Per_Store.ProductName = (string)rdrR["ProductName"];
        //                Match_SKU_Per_Store.OurPrice = (string)rdrR["OurPrice"];
        //                Match_SKU_Per_Store.ProductSKU = (string)rdrR["ProductSKU"];
        //                Match_SKU_Per_Store.UPC = (string)rdrR["UPC"];
        //                Match_SKU_Per_Store.ABFixtures = (string)rdrR["ABFixtures"];
        //                Match_SKU_Per_Store.AllenDisplay = (string)rdrR["AllenDisplay"];
        //                Match_SKU_Per_Store.Amazon = (string)rdrR["Amazon"];
        //                Match_SKU_Per_Store.AmazonAmazon = (string)rdrR["AmazonAmazon"];
        //                Match_SKU_Per_Store.AmericanRetailSupply = (string)rdrR["AmericanRetailSupply"];
        //                Match_SKU_Per_Store.BarrDisplay = (string)rdrR["BarrDisplay"];
        //                Match_SKU_Per_Store.Bellacor = (string)rdrR["Bellacor"];
        //                Match_SKU_Per_Store.Bonanza = (string)rdrR["Bonanza"];
        //                Match_SKU_Per_Store.BuyStoreFixtures = (string)rdrR["BuyStoreFixtures"];
        //                Match_SKU_Per_Store.DawsonJones = (string)rdrR["DawsonJones"];
        //                Match_SKU_Per_Store.DealerSupply = (string)rdrR["DealerSupply"];
        //                Match_SKU_Per_Store.DisplayWarehouse = (string)rdrR["DisplayWarehouse"];
        //                Match_SKU_Per_Store.eproducthunterAmazon = (string)rdrR["eproducthunterAmazon"];
        //                Match_SKU_Per_Store.GABPAuto = (string)rdrR["GABPAuto"];
        //                Match_SKU_Per_Store.GarmentRackStoreAmazon = (string)rdrR["GarmentRackStoreAmazon"];
        //                Match_SKU_Per_Store.GemsOnDisplay = (string)rdrR["GemsOnDisplay"];
        //                Match_SKU_Per_Store.Globalindustrial = (string)rdrR["Globalindustrial"];
        //                Match_SKU_Per_Store.GPPInc = (string)rdrR["GPPInc"];
        //                Match_SKU_Per_Store.Hangers = (string)rdrR["Hangers"];
        //                Match_SKU_Per_Store.HangersDirect = (string)rdrR["HangersDirect"];
        //                Match_SKU_Per_Store.JewelrySupply = (string)rdrR["JewelrySupply"];
        //                Match_SKU_Per_Store.KCStoreFixtures = (string)rdrR["KCStoreFixtures"];
        //                Match_SKU_Per_Store.MillerSupplyIncAmazon = (string)rdrR["MillerSupplyIncAmazon"];
        //                Match_SKU_Per_Store.MyDealerSupply = (string)rdrR["MyDealerSupply"];
        //                Match_SKU_Per_Store.NaHanCo = (string)rdrR["NaHanCo"];
        //                Match_SKU_Per_Store.NashvilleWraps = (string)rdrR["NashvilleWraps"];
        //                Match_SKU_Per_Store.OnlyGarmentRacks = (string)rdrR["OnlyGarmentRacks"];
        //                Match_SKU_Per_Store.OnlyHangers = (string)rdrR["OnlyHangers"];
        //                Match_SKU_Per_Store.Onlymannequins = (string)rdrR["Onlymannequins"];
        //                Match_SKU_Per_Store.PalayDisplay = (string)rdrR["PalayDisplay"];
        //                Match_SKU_Per_Store.PaperMart = (string)rdrR["PaperMart"];
        //                Match_SKU_Per_Store.ProductDisplaySolutions = (string)rdrR["ProductDisplaySolutions"];
        //                Match_SKU_Per_Store.RetailPackaging = (string)rdrR["RetailPackaging"];
        //                Match_SKU_Per_Store.RetailResource = (string)rdrR["RetailResource"];
        //                Match_SKU_Per_Store.RoxyDisplay = (string)rdrR["RoxyDisplay"];
        //                Match_SKU_Per_Store.SelbyStoreFixtures = (string)rdrR["SelbyStoreFixtures"];
        //                Match_SKU_Per_Store.SIDSavage = (string)rdrR["SIDSavage"];
        //                Match_SKU_Per_Store.SpecialtyStoreServices = (string)rdrR["SpecialtyStoreServices"];
        //                Match_SKU_Per_Store.SprinklesGiftsAmazon = (string)rdrR["SprinklesGiftsAmazon"];
        //                Match_SKU_Per_Store.StampsStoreFixtures = (string)rdrR["StampsStoreFixtures"];
        //                Match_SKU_Per_Store.Staples = (string)rdrR["Staples"];
        //                Match_SKU_Per_Store.StoreFixture = (string)rdrR["StoreFixture"];
        //                Match_SKU_Per_Store.TheFixtureZone = (string)rdrR["TheFixtureZone"];
        //                Match_SKU_Per_Store.TSISupplies = (string)rdrR["TSISupplies"];
        //                Match_SKU_Per_Store.ULine = (string)rdrR["ULine"];
        //                Match_SKU_Per_Store.Wawak = (string)rdrR["Wawak"];

        //                checker.Add(Match_SKU_Per_Store.BrandSKU, Match_SKU_Per_Store);

        //            }
        //            cont.Close();
        //            rdrR.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
        //    }
        //    finally
        //    {
        //        if (cont != null)
        //        {
        //            cont.Close();
        //        }
        //        if (rdrR != null)
        //        {
        //            rdrR.Close();
        //        }
        //    }
        //}
        //public void GetSKUandStore(Dictionary<string, SKUandStore> KeySKU_ValueSKUandStore_eachStore, string sql, string find) //Data for the above method
        //{
        //    SqlConnection cont = null;
        //    SqlDataReader rdrR = null;




        //    try
        //    {
        //        using (cont = new SqlConnection(CSPG))
        //        {
        //            cont.Open();
        //            SqlCommand cmd1 = new SqlCommand(sql, cont);




        //            cmd1.CommandType = CommandType.Text;

        //            rdrR = cmd1.ExecuteReader();

        //            while (rdrR.Read())
        //            {
        //                SKUandStore sKUandStorE = new SKUandStore();

        //                sKUandStorE.BrandSKU = (string)rdrR["BrandSKU"];
        //                sKUandStorE.Price = (string)rdrR[find];

        //                KeySKU_ValueSKUandStore_eachStore.Add(sKUandStorE.BrandSKU, sKUandStorE);
        //            }
        //            cont.Close();
        //            rdrR.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
        //    }
        //    finally
        //    {
        //        if (cont != null)
        //        {
        //            cont.Close();
        //        }
        //        if (rdrR != null)
        //        {
        //            rdrR.Close();
        //        }
        //    }
        //}

        //public void WriteToExcelNameWithOrderId(Dictionary<string, string> pairs)// Used to check things randomly
        //{
        //    Application ExcelApp = new Application();
        //    Workbook ExcelWorkBook = null;

        //    Worksheet ExcelWorkSheet = null;
        //    ExcelApp.Visible = true;

        //    ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

        //    try
        //    {
        //        int r = 1; int c = 1;

        //        ExcelWorkSheet = ExcelWorkBook.Worksheets[1];

        //        ExcelWorkSheet.Cells[r, c + 0] = "OrderId";
        //        ExcelWorkSheet.Cells[r, c + 1] = "StoreName";


        //        r++;
        //        string name = "";
        //        foreach (var item in pairs)
        //        {
        //            if (name == "")
        //            {
        //                name = item.Value;
        //            }
        //            if (name != item.Value)
        //            {
        //                r = 1;
        //                c += 2;
        //                name = item.Value;
        //            }
        //            ExcelWorkSheet.Cells[r, c + 0] = item.Key.Substring(0, 8);
        //            ExcelWorkSheet.Cells[r, c + 1] = item.Value;


        //            r++;

        //        }
        //        ExcelWorkBook.Worksheets[1].Name = "Before";
        //        ExcelWorkBook.SaveAs("d:\\Keys.xlsx");

        //        ExcelWorkBook.Close();
        //        ExcelApp.Quit();
        //        Marshal.ReleaseComObject(ExcelWorkSheet);
        //        Marshal.ReleaseComObject(ExcelWorkBook);
        //        Marshal.ReleaseComObject(ExcelApp);
        //    }

        //    catch (Exception ex)

        //    {

        //        WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace + ex.GetType());

        //    }

        //    finally

        //    {

        //        foreach (Process process in Process.GetProcessesByName("Excel"))

        //            process.Kill();

        //    }

        //}
    }
}