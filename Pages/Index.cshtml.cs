using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Newtonsoft.Json;
using System.Net.Http.Headers;

namespace ADOReportWeb.Pages
{
    public class IndexModel : PageModel
    {
        const string ADO_ORG_NAME = "";
        const string ADO_PROJECT_NAME = "";
        const string ADO_QUERY_ID = "";

        private readonly ILogger<IndexModel> _logger;

        public string HtmlData { get; set; }

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            try
            {
                var credential = new DefaultAzureCredential();

                // Create an instance of HttpClient.
                var client = new HttpClient();
                string AdoAppClientID = "499b84ac-1321-427f-aa17-267ca6975798/.default";
                string[] scopes = new string[] { AdoAppClientID };

                // Now get the the Access token
                var result = credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes)).Result;

                // Build the ADO REST API Urls
                client.BaseAddress = new Uri(string.Format("https://dev.azure.com/{0}/", ADO_ORG_NAME));  //url of your organization
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                Console.WriteLine(result.Token);
                // Assignt the access token
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.Token);

                // Call the ADO API to get the list of GitHub backlog items
                HttpResponseMessage response = client.GetAsync(string.Format("{0}/_apis/wit/wiql/{1}?api-version=5.1", ADO_PROJECT_NAME, ADO_QUERY_ID)).Result;
                if (response.IsSuccessStatusCode)
                {
                    // Get the response as a string
                    string rootData = response.Content.ReadAsStringAsync().Result;

                    // Convert the response to C# Object using Newtonsoft.JSON library
                    // This response will give us a list of all the GitHub Backlog work item Ids
                    Root root = JsonConvert.DeserializeObject<Root>(rootData);

                    // Now we need to do an other query to get the details that we need to each item
                    // For that query, we will need the comma separated list of the all item ids and the name of the fields that we need
                    string ids = "";
                    foreach (var item in root.workItems)
                    {
                        ids = string.IsNullOrEmpty(ids) ? ids + item.id.ToString() : ids + "," + item.id.ToString();
                    }

                    // Here are the name of the fields that we need
                    string postData = "{'ids': [" + ids + "],'fields': ['System.Id','System.Title','System.Description']}";
                    var content = new StringContent(postData);

                    // Build the request and call the REST API
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                    response = client.PostAsync("_apis/wit/workitemsbatch?api-version=7.1-preview.1", content).Result;

                    // Get the response as a string
                    string details = response.Content.ReadAsStringAsync().Result;

                    // Convert the response to C# Object using Newtonsoft.JSON library
                    // This response will give us the requested details about all the GitHub Backlog work item Ids
                    DetailedRoot detailedRoot = JsonConvert.DeserializeObject<DetailedRoot>(details);

                    // Now build the HTML with an JTML table

                    // Initial HTML tag and style
                    string htmlData = "<HTML><style>table, th, td {border: 1px solid black;border-collapse: collapse;}</style>";

                    // Add the P tag for Item count
                    htmlData = htmlData + "<p>Total Number of Items = " + detailedRoot.value.Count.ToString() + "</p>";

                    // HTML Table Headers
                    string htmlTable = "<table><tr><th>Id</th><th>Title</th><th>Description</th></tr>";

                    // Loop through each item to create an HTML table row for each
                    foreach (var item in detailedRoot.value)
                    {
                        // Build HTML Table Row
                        string htmlTableRow = string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>", item.fields.SystemId, item.fields.SystemTitle, item.fields.SystemDescription, item.fields.AcceptanceCriteria);
                        htmlTable = htmlTable + htmlTableRow;
                    }

                    // HTML table closing tag
                    htmlTable = htmlTable + "</table>";

                    // HTML close tag
                    htmlData = htmlData + htmlTable + "</HTML>";
                    HtmlData = htmlData;
                }
                else
                {
                    // Write error to console if the response is not 200
                    HtmlData = "An exception has occurred:" + response.StatusCode.ToString();
                    Console.WriteLine(response.StatusCode.ToString());
                }
            }
            catch (Exception ex)
            {
                HtmlData = "An exception has occurred:" + ex.Message;
            }

            Console.WriteLine(HtmlData);
        }
    }

    // All the classes below are generated using the online JSON to C# tools
    public class Column
    {
        public string referenceName { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }
    public class Field
    {
        public string referenceName { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }
    public class Root
    {
        public string queryType { get; set; }
        public string queryResultType { get; set; }
        public DateTime asOf { get; set; }
        public List<Column> columns { get; set; }
        public List<SortColumn> sortColumns { get; set; }
        public List<WorkItem> workItems { get; set; }
    }
    public class SortColumn
    {
        public Field field { get; set; }
        public bool descending { get; set; }
    }

    public class WorkItem
    {
        public int id { get; set; }
        public string url { get; set; }
    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class CommentVersionRef
    {
        public int commentId { get; set; }
        public int version { get; set; }
        public string url { get; set; }
    }

    public class Fields
    {
        [JsonProperty("System.Id")]
        public int SystemId { get; set; }

        [JsonProperty("System.Title")]
        public string SystemTitle { get; set; }

        [JsonProperty("System.Description")]
        public string SystemDescription { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.AcceptanceCriteria")]
        public string AcceptanceCriteria { get; set; }
    }

    public class DetailedRoot
    {
        public int count { get; set; }
        public List<Value> value { get; set; }
    }

    public class Value
    {
        public int id { get; set; }
        public int rev { get; set; }
        public Fields fields { get; set; }
        public string url { get; set; }
        public CommentVersionRef commentVersionRef { get; set; }
    }
}