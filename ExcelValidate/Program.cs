using OfficeOpenXml;
using System;
using System.IO;
using System.Net.Http;

namespace ExcelValidate
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //Excel file location
            var fi = new FileInfo(@"C:\Users\Rajesh\source\repos\ExcelValidate\ExcelValidate\ExcelData.xlsx");
            //Creating Exccel file package that will contain ExcelData.xlsx
            ExcelPackage excel = new ExcelPackage(fi);
            //Specifying sheet name
            ExcelWorksheet firstWorksheet = excel.Workbook.Worksheets["Sheet1"];
            //Creating for loop to print all the rows available in Sheet1
            //firstWorksheet.Dimension.End.Row gets the last row number that contains data
            for (int i = 1; i <= firstWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= 5; j++)
                {
                    Console.Write(firstWorksheet.Cells[i, j].Value + " ");
                    //writing data in excel
                    //firstWorksheet.Cells[i, 6].Value = "True";
                }
                Console.WriteLine("\n");
            }
            excel.Save();
        }
    }
}
