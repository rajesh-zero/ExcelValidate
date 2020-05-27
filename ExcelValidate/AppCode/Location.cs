class Location:BaseEntity
{
    string address, pincode;

    public string Address { get => address; set => address = value; }
    public string Pincode { get => pincode; set => pincode = value; }
}