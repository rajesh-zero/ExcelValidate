using OfficeOpenXml;
using RestSharp;
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
            var fi = new FileInfo(@"..\..\..\ExcelData.xlsx"); //relative path for excel file
            //Creating Exccel file package that will contain ExcelData.xlsx
            ExcelPackage excel = new ExcelPackage(fi);
            //Specifying sheet name
            ExcelWorksheet firstWorksheet = excel.Workbook.Worksheets["Sheet1"];
            //Creating for loop to print all the rows available in Sheet1
            //firstWorksheet.Dimension.End.Row gets the last row number that contains data
            for (int i = 2; i <= firstWorksheet.Dimension.End.Row; i++)
            {
                var Name = firstWorksheet.Cells[i, 1].Value;
                DateTime DateOfBirth;
                bool IsActive = (bool)firstWorksheet.Cells[i, 3].Value;
                var Balance = firstWorksheet.Cells[i, 4].Value;
                var LoanAmount = firstWorksheet.Cells[i, 5].Value;

                Console.WriteLine("results");
                Console.WriteLine(Name != null);
                Console.WriteLine(DateTime.TryParse(firstWorksheet.Cells[i, 2].Value.ToString(), out DateOfBirth) == true);
                Console.WriteLine(IsActive is true);
                Console.WriteLine(Balance != null);
                Console.WriteLine(LoanAmount != null);
                //Console.WriteLine(Name != null && DateTime.TryParse(firstWorksheet.Cells[i, 2].Value.ToString(), out DateOfBirth) == true && IsActive.ToString() == "true" && Balance != null && LoanAmount != null);
                if( Name != null 
                    && DateTime.TryParse(firstWorksheet.Cells[i, 2].Value.ToString(), out DateOfBirth) == true
                    && IsActive is true
                    && Balance != null 
                    && LoanAmount != null)
                {
                    Console.WriteLine(Name.ToString()
                        + DateOfBirth.ToString()
                        + IsActive.ToString()
                        + Balance.ToString()
                        + LoanAmount.ToString());
                }
                Console.WriteLine("\n");
            }
            excel.Save();
        }
        static void apiRequest(string name)
        {
            //future reference var param = new MyClass { IntData = 1, StringData = "test123" };

            //creating restclient with url to send post request
            var client = new RestClient("https://httpbin.org/");
            //specifying api resource path i.e https://httpbin.org/post and specifying request format as json
            var request = new RestRequest("post", DataFormat.Json);
            request.AddParameter("name",name, ParameterType.GetOrPost);
            //future reference request.AddJsonBody(param);
            var response = client.Post(request);
            //var response = client.Get(request);
            Console.WriteLine(response.Content);
        }
    }
}
