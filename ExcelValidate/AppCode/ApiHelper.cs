using System;
using System.Collections.Generic;
using System.Net;
using System.Reflection;
using Newtonsoft.Json.Linq;
using RestSharp;

class ApiHelper
{
    private RestClient client;
    private RestRequest request;
    public ApiHelper(string url)
    {
        client = new RestClient(url);
    }
    public void BuildRequest(string resource,Person p)
    {
        request = new RestRequest("post", DataFormat.Json);
        request.AddParameter("name",p.Name,ParameterType.GetOrPost);
        request.AddParameter("dob",p.DateOfBirth,ParameterType.GetOrPost);
        request.AddParameter("isActive",p.IsActive,ParameterType.GetOrPost);
        request.AddParameter("balannce",p.Balance,ParameterType.GetOrPost);
        request.AddParameter("loanAmount",p.LoanAmount,ParameterType.GetOrPost);
    }
    public bool SendPost()
    {
        bool returnStatus = false;
        var response = client.Post(request);
        if (response.StatusCode == HttpStatusCode.OK)
        {
            Console.WriteLine("Status is " + response.StatusCode);
            var jObject = JObject.Parse(response.Content);
            if (jObject["form"] != null)
            {
                returnStatus = true;
            }
            else
            {
                Console.WriteLine("Should not encounter this error ");
            }
        }
        else
        {
            Console.WriteLine("Status is " + response.StatusCode + "Failed");
        }
        return returnStatus;
    }
}