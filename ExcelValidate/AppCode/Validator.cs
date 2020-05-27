using System;
using System.Reflection;
using System.Text.RegularExpressions;

interface IValidator
{
    bool Validate(BaseEntity e);
}
public class Validator
{
    public static bool Validate(BaseEntity e)
    {
        IValidator iv;
        if (e is Person)
        {
            iv = new ValidatePerson();
            return iv.Validate(e);
        }
        else if (e is Location)
        {
            iv = new ValidateLocation();
            return iv.Validate(e);
        }
        return false;

    }
}

public class ValidateLocation : IValidator
{
    public bool Validate(BaseEntity e)
    {
        Console.WriteLine(e);
        return true;
    }
}

public class ValidatePerson : IValidator
{
    public bool Validate(BaseEntity e)
    {
        Person t = (Person)e;
        bool f, g, h, i, j, k;
        DateTime tempdt;
        f = Regex.IsMatch(t.Name, @"^[a-zA-Z]+$");
        g = DateTime.TryParse(t.DateOfBirth, out tempdt);
        h = Regex.IsMatch(t.IsActive, @"True|False");
        i = Regex.IsMatch(t.Balance.ToString(), @"^[0-9]+.\d{0,2}$");
        j = Regex.IsMatch(t.LoanAmount.ToString(), @"^[0-9]+.\d{0,2}$");
        k = f && g && h && i && j;
        if (k == true)
        {
            return true;
        }
        else
        {
            throw new InvalidDataException();
        }
    }
}