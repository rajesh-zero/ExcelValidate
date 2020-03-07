using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using RestSharp;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;

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
                DateTime tempdt; //tempdt temp variable to work with DateOfBirth column data
                //getting column values from excel in var for further processing of data
                var rawName = firstWorksheet.Cells[i, 1].Value;
                var rawDateOfBirth = firstWorksheet.Cells[i, 2].Value;
                var rawIsActive = firstWorksheet.Cells[i, 3].Value;
                var rawBalance = firstWorksheet.Cells[i, 4].Value;
                var rawLoanAmount = firstWorksheet.Cells[i, 5].Value;

                bool apiCheckStatus = false; //setting default value if api call fails or is not needed 

                //checking data for null values
                if (rawName != null && rawDateOfBirth != null && rawIsActive != null && rawBalance != null && rawLoanAmount != null)
                {
                    //creating boolean variables to validate if data is in desired format
                    bool validName = Regex.IsMatch(rawName.ToString(), @"^[a-zA-Z]+$"); //name takes only alphabets a-z and A-Z
                    bool validDateOfBirth = DateTime.TryParse(rawDateOfBirth.ToString(), out tempdt);
                    bool validIsActive = false; //default value if data from excel is not boolean
                    //if condition to check if rawIsActive contains boolean values if yes then assign the value to validIsActive
                    if (rawIsActive is bool)
                    {
                        validIsActive = (bool)rawIsActive;
                    }
                    bool validBalance = Regex.IsMatch(rawBalance.ToString(), @"^[0-9]+.\d{0,2}$"); //only takes numbers upto 2 decimal
                    bool validLoanAmount = Regex.IsMatch(rawLoanAmount.ToString(), @"^[0-9]+.\d{0,2}$"); //only takes numbers upto 2 decimal
                    //apiCheckStatus = "false";
                    if (validName is true && validDateOfBirth is true && validIsActive is true && validBalance is true && validLoanAmount is true)
                    {
                        Console.Write("Data is Valid..... Making api call....");
                        apiCheckStatus = ApiRequest(rawName.ToString(),rawDateOfBirth.ToString(), rawIsActive.ToString(),rawBalance.ToString(),rawLoanAmount.ToString());
                        
                        Console.WriteLine("Row " + i + "----" + rawName.ToString() + "----" +
                            rawDateOfBirth.ToString() + "----" +
                            rawIsActive.ToString() + "----" +
                            rawBalance.ToString() + "----" +
                            rawLoanAmount.ToString() + "---- update column as "+apiCheckStatus);
                    }
                    else
                    {
                        Console.WriteLine("Data Validation error on row " + i);
                    }
                }
                else
                {
                    Console.WriteLine("Empty values encountered on row " + i);
                }
                firstWorksheet.Cells[i, 6].Value = apiCheckStatus.ToString();
                Console.WriteLine("\n");
            }
            excel.Save();
        }
        static bool ApiRequest(string name,string dob,string isactive,string balance,string loanamount)
        {
            //future reference var param = new MyClass { IntData = 1, StringData = "test123" };

            //creating restclient with url to send post request
            var client = new RestClient("https://httpbin.org/");
            //specifying api resource path i.e https://httpbin.org/post and specifying request format as json
            var request = new RestRequest("post", DataFormat.Json);
            request.AddParameter("Name",name, ParameterType.GetOrPost);
            request.AddParameter("DateOfBirth",dob, ParameterType.GetOrPost);
            request.AddParameter("IsActive",isactive, ParameterType.GetOrPost);
            request.AddParameter("Balance",balance, ParameterType.GetOrPost);
            request.AddParameter("LoanAmount",loanamount, ParameterType.GetOrPost);
            //future reference request.AddJsonBody(param);
            var response = client.Post(request);
            //var response = client.Get(request);
            
            if(response.StatusCode == HttpStatusCode.OK)
            {
                Console.WriteLine("Status is " + response.StatusCode);
                var jObject = JObject.Parse(response.Content);
                //Console.WriteLine(jObject["form"]["Name"]);
                if(jObject["form"]["Name"].ToString() == name)
                {
                    return true;
                }
                else
                {
                    Console.WriteLine("unknown error line 101 ");
                }
            }
            else
            {
                Console.WriteLine("Status is " + response.StatusCode + "Failed");
                return false;
            }
            return false;
        }

    }
}
