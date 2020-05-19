using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.IO;

namespace ExcelValidate
{
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
