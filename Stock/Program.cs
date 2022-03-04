using System;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;

namespace Stock
{
    class Program
    {
        /*static float EPSGap(float esp1, float esp2)
        {
            return esp2 - esp1;
        }
        static void PERAve(float[] pers, out float perAve)
        {
            perAve = 0;

            foreach (var value in pers)
                perAve += value;
        }
        static float Power(int times, float value)
        {
            for (int i = 0; i < times; i++)
                value *= value;
            return value;
        }
        static float Sales(int times, float salePercent, float price)
        {
            for (int i = 0; i < times; i++)
                price /= salePercent;
            return price;
        }
        static void Main(string[] args)
        {
            float espGap = 0;
            float espAve = 0;
            float perAve = 0;
            float oriPrice = 0;
            float salePrice = 0;
            float[] pers = new float[5];

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(@"D:\Works\C#\AutoStock\A.xlsx");

            //Excel.Worksheet ws = wb.Worksheets.Item[1];

            //Console.WriteLine(ws.Cells[1, 1].value.GetType());
            //Console.WriteLine(ws.Cells[1, 1].value);
            try
            {
                for(int i = 1; i < 2; i++)  //스프레드시트 검색
                {
                     Excel.Worksheet ws = wb.Worksheets.Item[i];

                    for (int j = 1; j < 100; j += 2)    //주식수
                    {
                        if (ws.Cells[j, 1].value == null)
                            break;
                        for (int k = 3; k < 7; k++)  //esp 성장차이
                        {
                            espGap += EPSGap((float)ws.Cells[j, k].value, (float)ws.Cells[j, k + 1].value);
                        }
                        espAve = Power(5, espGap / 4);

                        for (int l = 0; l < 5; l++)
                        {
                            pers[l] = (float)ws.Cells[j + 1, l + 3].value;
                        }
                        PERAve(pers, out perAve);

                        oriPrice = perAve * espAve;
                        salePrice = Sales(5, 1.1f, oriPrice);

                        ws.Cells[j, 2] = salePrice;
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
            //ws.Cells[1, 2] = ws.Cells[1, 1];
            //Console.WriteLine(ws.Cells[1, 2].value);
            wb.SaveAs(@"D:\Works\C#\AutoStock\A.xlsx");
            wb.Close();
            excelApp.Quit()
            }
        }*/
        static List<string> GetStockNum()
        {
            bool end = false;
            int stockCount = 1;
            List<string> stockNumList = new List<string>();
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(@"D:\Works\C#\AutoStock\주식.xlsx");
            try
            {
                Excel.Worksheet ws = wb.Worksheets.Item[1];
                
                while(end == false)
                {
                    stockNumList.Add((string)ws.Cells[stockCount, 1].value);
                    Console.WriteLine(stockCount);
                    stockCount++;
                    
                    if (ws.Cells[stockCount, 1].value == null)
                        end = true;
                }
                Console.WriteLine(stockCount);

                foreach (string a in stockNumList)
                    Console.WriteLine(a);
                return stockNumList;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                return stockNumList;
            }
            finally
            {
                wb.Close();
                excelApp.Quit();
            }
        }
        static void GetESP(string a)
        {
            IWebDriver driver = new ChromeDriver();
            try
            {
                Console.WriteLine(String.Format("https://navercomp.wisereport.co.kr/v2/company/c1040001.aspx?cmp_cd={0}&cn=", a));
                driver.Url = String.Format("https://navercomp.wisereport.co.kr/v2/company/c1040001.aspx?cmp_cd={0}&cn=", a);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                Console.WriteLine("a");
                for (int i = 2; i < 7; i++)
                {
                    var table = driver.FindElement(By.XPath($"/html/body/div/form/div[1]/div/div[2]/div[3]/div/div/div[9]/table[2]/tbody/tr[1]/td[{i}]"));

                    Console.WriteLine(table.Text);
                }
                Console.WriteLine("b");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        static void Main(string[] args)
        {
            List<string> stocks = GetStockNum();

            foreach(string stock in stocks)
            {
                GetESP(stock);
            }
        }
        
    }
}
