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
                    //taking name column value in namevar
                    var namevar = firstWorksheet.Cells[i, j].Value;
                    if (j == 1 && namevar != null && i > 1)//if name is not null then
                    {
                        
                        apiRequest(firstWorksheet.Cells[i, j].Value.ToString()); //call apiRequestmethod
                        firstWorksheet.Cells[i, 6].Value = "Success"; //update success at row end
                    }
                    else if(j == 1 && namevar == null && i > 1)//if name is null then
                    {
                        firstWorksheet.Cells[i, 6].Value = "failure"; //update failure at row end
                    }
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
