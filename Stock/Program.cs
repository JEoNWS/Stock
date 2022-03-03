using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Stock
{
    class Program
    {
        static void Main(string[] args)
        {
            int espAve;
            int espGap;
            int perAve;
            int perSum;

            Excel.Application excelApp = new Excel.Application();   //엑셀 인스턴스 생성

            Excel.Workbook wb = excelApp.Workbooks.Open(@"D:\Works\C#\AutoStock\A.xlsx");  //엑셀 파일 불러오기

            for(int i = 1; i < 1; i++)
            {
                Console.WriteLine("1");
                espAve = 0;
                espGap = 0;
                perSum = 0;
                perAve = 0;

                Excel.Worksheet ws = wb.Worksheets.Item[i];
                
                for(int j = 0; j < 10; j++)
                {
                    //if (ws.Cells[j + 1, 1] == null)
                        //break;
                    Console.WriteLine("Start");
                    for(int k = 3; k < 7; k++)
                    {
                      espGap += ws.Cells[j + 1, k + 1] - ws.Cells[j + 1, k];
                    }
                    espAve = espGap / 4;
                    espAve = ws.Cells[j+1, 8] * (espAve * espAve * espAve * espAve * espAve);

                    for(int l = 8; l <= 12; l++)
                    {
                        perSum += ws.Cells[j + 1, l];
                    }
                    perAve = perSum / 5;

                    ws.Cells[j + 1, 2] = perAve * espAve;
                }
            }
            wb.Save();
            excelApp.Quit();
            Console.WriteLine("Done");
        }
    }
}
