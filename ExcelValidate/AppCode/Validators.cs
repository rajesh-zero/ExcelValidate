using System;
using System.Text.RegularExpressions;

interface IValidator<T>
{
    bool ValidatePerson(T t);
}

class Validator : IValidator<Person>
{

    public bool ValidatePerson(Person t)
    {
        // validation logic
        bool a, b, c, d, e, x;
        DateTime tempdt;
        a = Regex.IsMatch(t.Name, @"^[a-zA-Z]+$");
        b = DateTime.TryParse(t.DateOfBirth, out tempdt);
        c = Regex.IsMatch(t.IsActive, @"True|False");
        d = Regex.IsMatch(t.Balance.ToString(), @"^[0-9]+.\d{0,2}$");
        e = Regex.IsMatch(t.LoanAmount.ToString(), @"^[0-9]+.\d{0,2}$");
        x = a && b && c && d && e;
        if (x == true)
        {
            return true;
        }
        else
        {
            throw new InvalidDataException();
        }
    }

}