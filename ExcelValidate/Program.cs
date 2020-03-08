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

            //setting directory path because of https://github.com/dotnet/project-system/issues/3619
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);

            //Excel file location relative path
            var fi = new FileInfo(@"..\..\..\ExcelData.xlsx");

            //Creating Exccel file package that will contain ExcelData.xlsx
            ExcelPackage excel = new ExcelPackage(fi);

            //Specifying sheet name
            ExcelWorksheet firstWorksheet = excel.Workbook.Worksheets["Sheet1"];

            //Creating for loop to access all the rows available in Sheet1
            //firstWorksheet.Dimension.End.Row gets the last row number that contains data
            //Data will be accessed from 2nd row as 1st row will contain headings
            for (int i = 2; i <= firstWorksheet.Dimension.End.Row; i++)
            {
                //tempdt temp variable to work with DateOfBirth column data
                DateTime tempdt;

                //getting raw column values from excel in var for further processing
                var rawName = firstWorksheet.Cells[i, 1].Value;
                var rawDateOfBirth = firstWorksheet.Cells[i, 2].Value;
                var rawIsActive = firstWorksheet.Cells[i, 3].Value;
                var rawBalance = firstWorksheet.Cells[i, 4].Value;
                var rawLoanAmount = firstWorksheet.Cells[i, 5].Value;

                //setting default value if api call fails or is not needed 
                bool apiCheckStatus = false;

                //checking data for null values
                if (rawName != null && rawDateOfBirth != null && rawIsActive != null && rawBalance != null && rawLoanAmount != null)
                {
                    //creating boolean variables to validate if data is in desired format
                    bool validName, validDateOfBirth, validIsActive, validBalance, validLoanAmount;

                    //name takes only alphabets a-z and A-Z
                    validName = Regex.IsMatch(rawName.ToString(), @"^[a-zA-Z]+$");

                    //Checking if data is of datetype
                    validDateOfBirth = DateTime.TryParse(rawDateOfBirth.ToString(), out tempdt);

                    //default value if data from excel is not boolean
                    validIsActive = false;

                    //if condition to check if rawIsActive contains boolean values if yes then assign the value to validIsActive
                    if (rawIsActive is bool)
                    {
                        validIsActive = (bool)rawIsActive;
                    }

                    //only takes numbers upto 2 decimal in Balance 
                    validBalance = Regex.IsMatch(rawBalance.ToString(), @"^[0-9]+.\d{0,2}$");

                    //only takes numbers upto 2 decimal in LoanAmount
                    validLoanAmount = Regex.IsMatch(rawLoanAmount.ToString(), @"^[0-9]+.\d{0,2}$");

                    if (validName is true && validDateOfBirth is true && validIsActive is true && validBalance is true && validLoanAmount is true)
                    {
                        Console.WriteLine("Row " + i + "----" + rawName.ToString() + "----" +
                            rawDateOfBirth.ToString() + "----" +
                            rawIsActive.ToString() + "----" +
                            rawBalance.ToString() + "----" +
                            rawLoanAmount.ToString());

                        Console.Write("Data is Valid..... Making api call....");
                        //calling ApiRequest method
                        apiCheckStatus = ApiRequest(rawName.ToString(), rawDateOfBirth.ToString(), rawIsActive.ToString(), rawBalance.ToString(), rawLoanAmount.ToString());
                        Console.WriteLine(" updating column as " + apiCheckStatus);

                    }
                    else
                    {
                        Console.WriteLine("Row {0} Data Validation occured updating column as {1}", i, apiCheckStatus);
                    }
                }
                else
                {
                    Console.WriteLine("Row {0} Empty values encountered updating column as {1}", i, apiCheckStatus);
                }
                firstWorksheet.Cells[i, 6].Value = apiCheckStatus.ToString(); //updating success or failure status in excel
                Console.WriteLine("\n");
            }
            excel.Save();
        }
        static bool ApiRequest(string name, string dob, string isactive, string balance, string loanamount)
        {
            //future reference var param = new MyClass { IntData = 1, StringData = "test123" };

            //creating restclient with url to send post request
            var client = new RestClient("https://httpbin.org/");

            //making request format by specifying api resource path i.e https://httpbin.org/post and specifying request format as json and adding parameters
            var request = new RestRequest("post", DataFormat.Json);
            request.AddParameter("Name", name, ParameterType.GetOrPost);
            request.AddParameter("DateOfBirth", dob, ParameterType.GetOrPost);
            request.AddParameter("IsActive", isactive, ParameterType.GetOrPost);
            request.AddParameter("Balance", balance, ParameterType.GetOrPost);
            request.AddParameter("LoanAmount", loanamount, ParameterType.GetOrPost);

            //Executing request to get Response
            var response = client.Post(request);

            //Checking valid response from server
            //Will check if internet connection is working and host is reachable
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Console.Write("Status is " + response.StatusCode);
                var jObject = JObject.Parse(response.Content);

                //Comparing received form data
                if (jObject["form"]["Name"].ToString() == name)
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
