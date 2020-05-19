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
    class ExcelFile
    {
        public ExcelPackage excel;
        public ExcelWorksheet firstWorksheet;
        public ExcelFile(FileInfo fi)
        {
            excel = new ExcelPackage(fi);
            firstWorksheet = excel.Workbook.Worksheets["Sheet1"];
        }

        public void CheckFields(int i)
        {
            /*
            This method checks individual row that matches i
            */
                try
                {
                    string Name = firstWorksheet.Cells[i, 1].Value.ToString();
                    string DateOfBirth = firstWorksheet.Cells[i, 2].Value.ToString();
                    string IsActive = firstWorksheet.Cells[i, 3].Value.ToString();
                    string Balance = firstWorksheet.Cells[i, 4].Value.ToString();
                    string LoanAmount = firstWorksheet.Cells[i, 5].Value.ToString();
                    bool stats = ValidateField( Name, DateOfBirth, IsActive, Balance, LoanAmount);
                    bool apistatus = false;
                    if (stats == true)
                    {
                        apistatus = CallApi(Name, DateOfBirth, IsActive, Balance, LoanAmount);
                        Console.WriteLine("Updating column with "+apistatus);
                        firstWorksheet.Cells[i, 6].Value = apistatus.ToString();
                    }
                    else
                    {
                        Console.WriteLine("validation Failed");
                        Console.WriteLine("Updating column with "+apistatus);
                        firstWorksheet.Cells[i, 6].Value = apistatus.ToString();
                    }
                }
                catch (System.Exception)
                {                    
                    Console.WriteLine("Empty values detected at row {0}",i);
                    Console.WriteLine("Updating column with False");
                    firstWorksheet.Cells[i, 6].Value = "False".ToString();
                }
            excel.Save();
        }
   
        public void CheckFields()
        {
            /*
            This method checks all rows
            */

            for (int i = 2; i <= firstWorksheet.Dimension.End.Row; i++)
            {
                try
                {
                    string Name = firstWorksheet.Cells[i, 1].Value.ToString();
                    string DateOfBirth = firstWorksheet.Cells[i, 2].Value.ToString();
                    string IsActive = firstWorksheet.Cells[i, 3].Value.ToString();
                    string Balance = firstWorksheet.Cells[i, 4].Value.ToString();
                    string LoanAmount = firstWorksheet.Cells[i, 5].Value.ToString();
                    bool stats = ValidateField( Name, DateOfBirth, IsActive, Balance, LoanAmount);
                    bool apistatus = false;
                    if (stats == true)
                    {
                        apistatus = CallApi(Name, DateOfBirth, IsActive, Balance, LoanAmount);
                        Console.WriteLine("Updating column with "+apistatus);
                        firstWorksheet.Cells[i, 6].Value = apistatus.ToString();
                    }
                    else
                    {
                        Console.WriteLine("validation Failed");
                        Console.WriteLine("Updating column with "+apistatus);
                        firstWorksheet.Cells[i, 6].Value = apistatus.ToString();
                    }
                }
                catch (System.Exception)
                {                    
                    Console.WriteLine("Empty values detected at row {0}",i);
                    Console.WriteLine("Updating column with False");
                    firstWorksheet.Cells[i, 6].Value = "False".ToString();
                }
            }
            excel.Save();
        }
        private bool ValidateField(string name,string dob,string isactive,string balance,string loanamount)
        {
            /*
            This method checks if all column has data in expected format
            */
            bool a,b,c,d,e;
            DateTime tempdt;
            a = Regex.IsMatch(name, @"^[a-zA-Z]+$");
            b = DateTime.TryParse(dob, out tempdt);
            c = Regex.IsMatch(isactive, @"True|False");
            d = Regex.IsMatch(balance.ToString(), @"^[0-9]+.\d{0,2}$");
            e = Regex.IsMatch(loanamount.ToString(), @"^[0-9]+.\d{0,2}$");
            
            return a&&b&&c&&d&&e;
        }

        private bool CallApi(string name,string dob,string isactive,string balance,string loanamount)
        {
            /*
            This method takes parameters and makes a call to https://httpbin.org/
            If it gets same data in response then returns true for this records
            */
            bool returnStatus=false;
            var client = new RestClient("https://httpbin.org/");
            RestRequest request = new RestRequest("post", DataFormat.Json);
            request.AddParameter("Name", name, ParameterType.GetOrPost);
            request.AddParameter("DateOfBirth", dob, ParameterType.GetOrPost);
            request.AddParameter("IsActive", isactive, ParameterType.GetOrPost);
            request.AddParameter("Balance", balance, ParameterType.GetOrPost);
            request.AddParameter("LoanAmount", loanamount, ParameterType.GetOrPost);

            var response = client.Post(request);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                Console.WriteLine("Status is " + response.StatusCode);
                var jObject = JObject.Parse(response.Content);

                //Comparing received form data
                if (jObject["form"]["Name"].ToString() == name)
                {
                    returnStatus = true;
                }
                else
                {
                    Console.WriteLine("unknown error line 101 ");
                }
            }
            else
            {
                Console.WriteLine("Status is " + response.StatusCode + "Failed");
                returnStatus = false;
            }
            return returnStatus;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("-----------ExcelValidate-----------");
            //setting directory path because of https://github.com/dotnet/project-system/issues/3619
            //Very Important 
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
            FileInfo fi = new FileInfo(@"..\..\..\ExcelData.xlsx");
            ExcelFile e = new ExcelFile(fi);
            e.CheckFields();
            //e.CheckFields(17);
        }

    }
}
