using System;
using System.Collections.Generic;
using System.Net;
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
    public void FormRequest(Dictionary<string, string> PersonFields)
    {
        request = new RestRequest("post", DataFormat.Json);
        foreach (KeyValuePair<string, string> item in PersonFields)
        {
            string key = item.Key.ToString();
            string value = item.Value.ToString();
            request.AddParameter(key, value, ParameterType.GetOrPost);
        }
        //var response = client.Post(request);
    }
    public bool SendRequest()
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
            returnStatus = false;
        }
        return returnStatus;
    }
}