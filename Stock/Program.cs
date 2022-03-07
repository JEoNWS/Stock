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
        }*/

        /* static float Sales(int times, float salePercent, float price)
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
        static float Power(int times, float value)
        {
            float firstValue = value;
            for (int i = 0; i < times - 1; i++)
            {
                value *= firstValue;
            }
            return value;
        }
        static List<string> GetStockNum()
        {
            //bool end = false;
            int stockCount = 1;
            List<string> stockNumList = new List<string>();
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(@"D:\Works\C#\AutoStock\주식.xlsx");
            try
            {
                Excel.Worksheet ws = wb.Worksheets.Item[1];
                Excel.Range range = ws.UsedRange;

                for(int i = 1; i <= range.Rows.Count; i++)
                {
                    Console.WriteLine(range.Rows.Count);
                    stockNumList.Add(range.Cells[i, 2].Value);
                    Console.WriteLine(stockCount++);
                }
                
                /*while(end == false)
                {
                    stockNumList.Add(ws.Cells[stockCount, 2].Value);
                    Console.WriteLine(stockCount);
                    stockCount++;
                    

                    if (ws.Cells[stockCount, 2].Value == null)
                    {
                        Console.WriteLine("end");
                        end = true;
                    }
                }*/
                //Console.WriteLine(stockCount);

                //foreach (string a in stockNumList)
                    //Console.WriteLine(a);
                return stockNumList;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                return stockNumList;
            }
            finally
            {
                wb.Close(true);
                excelApp.Quit();
            }
        }
        static List<float> GetEPS(IWebDriver driver, string a, out string stockName)
        {
            List<float> EPSes = new List<float>();
            //IWebDriver driver = new ChromeDriver();
            try
            {
                //Console.WriteLine(String.Format("https://navercomp.wisereport.co.kr/v2/company/c1040001.aspx?cmp_cd={0}&cn=", a));
                driver.Url = String.Format("https://navercomp.wisereport.co.kr/v2/company/c1040001.aspx?cmp_cd={0}&cn=", a);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);  //창을 로드하기위해 기다리는 시간 로드전에 값을 불러올 수 없음
                //Console.WriteLine("a");
                stockName = driver.FindElement(By.XPath("/html/body/div/form/div[1]/div/div[2]/div[1]/div/table/tbody/tr[1]/td /dl/dt[1]/span")).Text;
                for (int i = 2; i < 7; i++)
                {
                    var eps = driver.FindElement(By.XPath($"/html/body/div/form/div[1]/div/div[2]/div[3]/div/div/div[9]/table[2]/tbody/tr[1]/td[{i}]"));

                    EPSes.Add(float.Parse(eps.Text));
                }
                //Console.WriteLine("b");
                return EPSes;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                stockName = "error";
                return EPSes;
            }
        }
        static List<float> GetPER(IWebDriver driver, string a, out string stockName)
        {
            List<float> PERs = new List<float>();

            try
            {
                float floatPer = 0f;
                driver.Url = String.Format("https://navercomp.wisereport.co.kr/v2/company/c1040001.aspx?cmp_cd={0}&cn=", a);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                //Console.WriteLine("c");
                stockName = driver.FindElement(By.XPath("/html/body/div/form/div[1]/div/div[2]/div[1]/div/table/tbody/tr[1]/td /dl/dt[1]/span")).Text;
                for(int i = 2; i < 7; i++)
                {
                    var per = driver.FindElement(By.XPath($"/html/body/div/form/div[1]/div/div[2]/div[3]/div/div/div[9]/table[2]/tbody/tr[17]/td[{i}]"));

                    if (float.TryParse(per.Text, out floatPer))
                        PERs.Add(floatPer);
                    else
                        PERs.Add(0f);
                }
                //Console.WriteLine("d");
                return PERs;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                stockName = "error";
                return PERs;
            }
        }
        static float CalcPrice(List<float> EPSes, List<float> PERs)
        {
            float[] EPSGap = new float[4];
            float aveEPS = 0;
            float sumPER = 0;
            float avePER = 0;
            float goalPrice = 0;
            float buyPrice = 0;
            float wantedMargin = 1.1f;

            for (int i = 0; i < EPSes.Count - 1; i++)
            {
                EPSGap[i] += (EPSes[i + 1] - EPSes[i]);
                EPSGap[i] /= Math.Abs(EPSes[i]);
                //Console.WriteLine(EPSGap[i]);
            }
            foreach (float epsGapPercent in EPSGap)
                aveEPS += epsGapPercent;
            aveEPS /= 4;
            aveEPS += 1;

            //Console.WriteLine(aveEPS);

            foreach (float PER in PERs)
                sumPER += PER;
            avePER = sumPER / 5;
            //Console.WriteLine(avePER);

            goalPrice = avePER * (EPSes[^1] * (Power(5, aveEPS)));
            //Console.WriteLine(goalPrice);
            buyPrice = (goalPrice / Power(5, wantedMargin));
            Console.WriteLine(buyPrice);

            return buyPrice;
        }
        static void WriteExcel(Excel.Worksheet ws, int index, string stockName, int price)
        {
            ws.Cells[index, 1] = stockName;
            /*if (price <= 0)
                ws.Cells[index, 2] = "에러";
            else*/
                ws.Cells[index, 2] = price.ToString();
            ws.Cells[index, 2].NumberFormat = "₩#,##0";
        }
        static void Main(string[] args)
        {
            int index = 1;
            float price = 0;
            int iPrice = 0;
            string stockName = "";
            List<string> stocks = GetStockNum();
            //Console.WriteLine(stocks.Count);

            IWebDriver driver = new ChromeDriver();

            List<float> EPSes = new List<float>();
            List<float> PERs = new List<float>();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(@"D:\Works\C#\AutoStock\주식결과.xlsx");
            Excel.Worksheet ws = wb.Worksheets.Item[1];

            foreach (string stock in stocks)
            {
                EPSes = GetEPS(driver, stock, out stockName);
                //foreach (float eps in EPSes)
                    //Console.WriteLine(eps);
                PERs = GetPER(driver, stock, out stockName);
                //foreach (float per in PERs)
                    //Console.WriteLine(per);
                price = CalcPrice(EPSes, PERs);
                //Console.WriteLine(price);
                iPrice = (int)price;
                WriteExcel(ws, index, stockName, iPrice);
                index++;
            }
            driver.Quit();
            wb.Save();
            wb.Close(true);
            excelApp.Quit();
        }
        
    }
}
