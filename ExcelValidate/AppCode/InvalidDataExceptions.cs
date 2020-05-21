using System;

class InvalidDataException : Exception
{
    public InvalidDataException()
    {

    }
    public InvalidDataException(string data) 
    : base(String.Format("Invalid Field data: {0}", data))
    {

    }
}