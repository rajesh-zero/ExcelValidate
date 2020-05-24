using System;

class InvalidDataException : Exception
{
    public InvalidDataException() 
    : base(String.Format("Invalid Field data: "))
    {

    }
}