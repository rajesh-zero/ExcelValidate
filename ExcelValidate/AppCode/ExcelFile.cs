using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
class ExcelFile
{
    public ExcelPackage excel;
    public ExcelWorksheet firstWorksheet;
    public ExcelFile(FileInfo fi)
    {
        excel = new ExcelPackage(fi);
        firstWorksheet = excel.Workbook.Worksheets["Sheet1"];
        Console.WriteLine("There are {0} records in this excel sheet", firstWorksheet.Dimension.End.Row);
    }
    public void ProcessRow(int i)
    {
        /*
        This method checks individual row that matches i
        */
        bool updateColumn = false;
        try
        {
            Person p = new Person();
            p.Name = firstWorksheet.Cells[i, 1].Value.ToString();
            p.DateOfBirth = firstWorksheet.Cells[i, 2].Value.ToString();
            p.IsActive = firstWorksheet.Cells[i, 3].Value.ToString();
            p.Balance = firstWorksheet.Cells[i, 4].Value.ToString();
            p.LoanAmount = firstWorksheet.Cells[i, 5].Value.ToString();

            ApiHelper a = new ApiHelper("https://httpbin.org/");
            a.BuildRequest("post", p);
            bool apistatus = a.SendPost();
            if (apistatus == true)
            {
                updateColumn = true;
            }
        }
        catch (System.NullReferenceException)
        {
            Console.WriteLine("Empty values");
        }
        catch (InvalidDataException e)
        {
            Console.WriteLine(e.Message);
        }
        UpdateResultInExcel(i, updateColumn);
        Console.WriteLine();
    }
    public void ProcessRows()
    {
        /*
        This method checks all rows
        */

        for (int i = 2; i <= firstWorksheet.Dimension.End.Row; i++)
        {
            ProcessRow(i);
        }
        excel.Save();
    }
    public string UpdateResultInExcel(int row, bool result)
    {
        Console.WriteLine("updating column as {0}", result);
        string updatestatus;
        try
        {
            firstWorksheet.Cells[row, 6].Value = result.ToString();
            excel.Save();
            updatestatus = "Success";
        }
        catch (System.Exception)
        {
            updatestatus = "Failed";
        }
        return updatestatus;
    }
}