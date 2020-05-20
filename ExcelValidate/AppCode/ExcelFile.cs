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

    public void CheckFields(int i)
    {
        /*
        This method checks individual row that matches i
        */
        bool updateColumn = false;
        Dictionary<string, string> PersonFields = new Dictionary<string, string>();
        try
        {
            PersonFields.Add("name", firstWorksheet.Cells[i, 1].Value.ToString());
            PersonFields.Add("dob", firstWorksheet.Cells[i, 2].Value.ToString());
            PersonFields.Add("isactive", firstWorksheet.Cells[i, 3].Value.ToString());
            PersonFields.Add("balance", firstWorksheet.Cells[i, 4].Value.ToString());
            PersonFields.Add("loanamount", firstWorksheet.Cells[i, 5].Value.ToString());
            Person p = new Person(PersonFields);
            bool validateResult = p.ValidateField();
            if (validateResult == true)
            {
                ApiHelper a = new ApiHelper("https://httpbin.org/");
                a.FormRequest(PersonFields);
                bool apistatus = a.SendRequest();
                if (apistatus == true)
                {
                    updateColumn = true;
                }
                else
                {
                    updateColumn = false;
                }
            }
            else
            {
                Console.WriteLine("Data validation error in row {0}", i);
                updateColumn = false;
            }
        }
        catch (System.Exception)
        {
            Console.WriteLine("Empty values detected at row {0}", i);
            updateColumn = false;
        }
        UpdateResult(i, updateColumn);
        Console.WriteLine();
    }

    public void CheckFields()
    {
        /*
        This method checks all rows
        */

        for (int i = 2; i <= firstWorksheet.Dimension.End.Row; i++)
        {
            CheckFields(i);
        }
        excel.Save();
    }
    public string UpdateResult(int row, bool result)
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
            updatestatus = "Faliled";
        }
        return updatestatus;
    }
}