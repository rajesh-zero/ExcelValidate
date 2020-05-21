using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

class Person
{

    private string name, dateOfBirth, isActive, balance, loanAmount;

    public string Name
    {
        get { return name; }

        set
        {
            if (Regex.IsMatch(value, @"^[a-zA-Z]+$") == true)
            {
                name = value;
            }
            else
            {
                throw new InvalidDataException(value);
            }
        }
    }
    public string DateOfBirth
    {
        get { return dateOfBirth; }
        set
        {
            DateTime tempdt;
            if ((DateTime.TryParse(value, out tempdt)) == true)
            {
                dateOfBirth = value;
            }
            else
            {
                throw new InvalidDataException(value);
            }
        }
    }
    public string IsActive
    {
        get { return isActive; }
        set
        {
            if ((Regex.IsMatch(value, @"True|False")) == true)
            {
                isActive = value;
            }
            else
            {
                throw new InvalidDataException(value);
            }
        }
    }
    public string Balance
    {
        get { return balance; }
        set
        {
            if (Regex.IsMatch(value.ToString(), @"^[0-9]+.\d{0,2}$") == true)
            {
                balance = value;
            }
            else
            {
                throw new InvalidDataException(value);
            }
        }
    }
    public string LoanAmount
    {
        get { return loanAmount; }
        set
        {
            if (Regex.IsMatch(value.ToString(), @"^[0-9]+.\d{0,2}$") == true)
            {
                loanAmount = value;
            }
            else
            {
                throw new InvalidDataException(value);
            }
        }
    }

}

