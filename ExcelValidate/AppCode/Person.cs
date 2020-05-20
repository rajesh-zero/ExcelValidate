using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

class Person
{
    private string Name,DateOfBirth,IsActive,Balance,LoanAmount;
    public Person(Dictionary<string, string> PersonFields)
    {
        Name = PersonFields["name"];
        DateOfBirth = PersonFields["dob"];
        IsActive = PersonFields["isactive"];
        Balance = PersonFields["balance"];
        LoanAmount = PersonFields["loanamount"];
    }

    public bool ValidateField()
    {
        /*
            This method checks if all column has data in expected format
            */
        bool a,b,c,d,e;
        DateTime tempdt;
        a = Regex.IsMatch(Name, @"^[a-zA-Z]+$");
        b = DateTime.TryParse(DateOfBirth, out tempdt);
        c = Regex.IsMatch(IsActive, @"True|False");
        d = Regex.IsMatch(Balance.ToString(), @"^[0-9]+.\d{0,2}$");
        e = Regex.IsMatch(LoanAmount.ToString(), @"^[0-9]+.\d{0,2}$");
            
        return a&&b&&c&&d&&e;
    }
}