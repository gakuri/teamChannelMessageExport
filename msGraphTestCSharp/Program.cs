// See https://aka.ms/new-console-template for more information
using System.Net;



main();

static void main()
{
    string AccessToken = Environment.GetCommandLineArgs()[1];
    string teamId=Environment.GetCommandLineArgs()[2];
    string channelId=Environment.GetCommandLineArgs()[3];
    string endPointOfMessages = string.Format("https://graph.microsoft.com/beta/teams/{0}/channels/{1}/messages", teamId, channelId);

    bool Continue = true;
    List<Newtonsoft.Json.Linq.JToken> messageList = new List<Newtonsoft.Json.Linq.JToken>();

    string responseBody;
    dynamic responseJson;

    while (Continue){
        responseBody = GetResponse(endPointOfMessages, AccessToken);
        responseJson = Newtonsoft.Json.JsonConvert.DeserializeObject(responseBody);

        foreach (Newtonsoft.Json.Linq.JToken message in responseJson.value)
        {
            messageList.Add(message);
        }
        if (responseJson["@odata.nextLink"] != null)
        {
            endPointOfMessages = responseJson["@odata.nextLink"];
        }
        else
        {
            Continue = false;
        }
    }

    string endPointOfReplies;
    List<Newtonsoft.Json.Linq.JToken> replyList;
    foreach (Newtonsoft.Json.Linq.JToken message in messageList)
    {
        replyList = new List<Newtonsoft.Json.Linq.JToken>();
        endPointOfReplies = string.Format("https://graph.microsoft.com/beta/teams/{0}/channels/{1}/messages/{2}/replies", teamId, channelId, (string)message["id"]);
        bool Continue2 = true;
        while (Continue2)
        {
            responseBody = GetResponse(endPointOfReplies, AccessToken);
            responseJson = Newtonsoft.Json.JsonConvert.DeserializeObject(responseBody);
            foreach (Newtonsoft.Json.Linq.JToken reply in responseJson.value)
            {
                replyList.Add(reply);
            }

            if (responseJson["@odata.nextLink"] != null)
            {
                endPointOfReplies = responseJson["@odata.nextLink"];
            }
            else
            {
                Continue2 = false;
                replyList.Add(message);
                System.IO.File.WriteAllText(@".\" + message["id"] + ".json", Newtonsoft.Json.JsonConvert.SerializeObject(replyList, Newtonsoft.Json.Formatting.Indented));
            }
        }


    }
}

static string GetResponse(string endpoint, string AccessToken)
{
    WebRequest Request = WebRequest.Create(endpoint);
    Request.Headers.Add("Authorization", AccessToken);
    string ResponseBody = "";

    using(WebResponse Response = Request.GetResponse())
    {
        using(System.IO.Stream Stream = Response.GetResponseStream())
        {
            using(StreamReader Reader = new StreamReader(Stream))
            {
                ResponseBody = Reader.ReadToEnd();
            }
        }
    }

    return ResponseBody;
}

