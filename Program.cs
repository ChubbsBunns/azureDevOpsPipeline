

using System.Net.Http.Headers;
using System.Text.Json;
using System.Net.Mail;
using SendGrid;
using SendGrid.Helpers.Mail;

using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using SendGrid.Extensions.DependencyInjection;
using System.Net.NetworkInformation;

const string USER_NEED_WORK_ITEM = "User Need";
const string DESIGN_INPUT_WORK_ITEM = "Design Input";
const string SRS_WORK_ITEM = "SRS";
const string TASK_WORK_ITEM = "Task";

const string JSON_DATA_VALUES = "value";

const string ORGANISATION_NAME = "OrganisationName:";
const string PROJECT_NAME = "ProjectName:";
const string PERSONAL_ACCESS_TOKEN = "PersonalAccessToken:";

var rawDataFromProject = "";
string ERROR_FILE_PATH = "error.txt";

string organisationName = "";
string projectName = "";
string personalAccessToken = "";

try
	{
        
        //getUserInputs(ref organisationName, ref projectName, ref personalAccessToken);
        string[] arguments = Environment.GetCommandLineArgs(); 
        getUserInputs(ref organisationName, ref projectName, ref personalAccessToken, arguments);
        Console.WriteLine("Organization name input is: " + organisationName);
        Console.WriteLine("Project name input is: " + projectName);
        Console.WriteLine("PAC input is: " + personalAccessToken);
        string getRequestUrl = parseGetRequestURL(projectName, organisationName);
        rawDataFromProject = getAzureDevOpsData(personalAccessToken, getRequestUrl).GetAwaiter().GetResult();
	}
	catch (Exception ex)
	{
        using (StreamWriter writer = new StreamWriter(ERROR_FILE_PATH))
        {
            writer.WriteLine("There was an error in getting information. It might be either a typo, or a get request issue from the server");
        }
        return;
	}

    //Data structures used

    //A lookup table of each work items id with their children work item ids.
    Dictionary<String, List<String>> workItemsData = new Dictionary<String, List<String>>();

    //A lookup table of each work item id with their corresponding work item type.
    Dictionary<String, String> workItemTypeLookup = new Dictionary<string, string>();

    //A lookup table of each work item id with whether they have been added to the table or not.
    Dictionary<String, bool> workItemAdded = new Dictionary<string, bool>();

    //A 2D array containing the data to be appended to the excel file.
    String[][] excelData = [];

    //An iterable stack that helps with populating each row in the excel file.
    List<String> workItemIdsToAdd = [];

    populateDefaultVariables(ref workItemsData, ref workItemTypeLookup, ref workItemAdded, ref excelData, ref rawDataFromProject);
    organiseExcelSheetData(workItemTypeLookup, workItemsData, ref workItemAdded, ref workItemIdsToAdd, ref excelData);
    
    string fileName = organisationName + "_" + projectName + "_workItem_AdjacencyMatrix.csv";
    // Write the data to the CSV file
    writeCsv(fileName, excelData);
    Console.WriteLine("CSV file created successfully at: " + Path.GetFullPath(fileName));

    /*
    Console.WriteLine("Trying to send email");
    try {
        Execute().Wait();
    } catch (Exception ex) {
        Console.WriteLine(ex);
    }
    */
    Console.WriteLine("Email is supposed to be sent"); 
    



//Helper functions

// Edits the organization name input, project input and PAC input
static void getUserInputs(ref string organisationName, ref string projectName, ref string personalAccessToken, string[] inputs) {
    string objectToAppendTo = "";
    foreach (string input in inputs) {
        if (input == ORGANISATION_NAME) {
            objectToAppendTo = ORGANISATION_NAME;
            continue;
        } else if (input == PROJECT_NAME) {
            objectToAppendTo = PROJECT_NAME;
            continue;
        } else if (input == PERSONAL_ACCESS_TOKEN) {
            objectToAppendTo = PERSONAL_ACCESS_TOKEN;
            continue;
        }

        switch(objectToAppendTo){
            case ORGANISATION_NAME:
                if (organisationName != "") {
                    organisationName = organisationName + " " + input;
                } else {
                    organisationName = input;
                }
                break;
            case PROJECT_NAME:
                if (projectName != "") {
                    projectName = projectName + " " + input;
                } else {
                    projectName = input;
                }

                break;
            case PERSONAL_ACCESS_TOKEN:
                personalAccessToken = personalAccessToken + input;
                break;
        }

    }
}

// Parses the user's inputs into a getrequest URL
static string parseGetRequestURL(string rawProjectName, string organizationName) {
    string parsedProjectTitle = "";
    foreach (char c in rawProjectName) {
        //The url requires white space to be replaced with %20
        if (c==' ') {
            parsedProjectTitle =  parsedProjectTitle + "%20";
        } else {
            parsedProjectTitle = parsedProjectTitle + c.ToString();
        }
    }
    return "https://analytics.dev.azure.com/" + organizationName + "/" + parsedProjectTitle + "/_odata/v4.0-preview/WorkItems?$select=WorkItemId, Title, WorkItemType, State&$expand=Children($select=WorkItemId,Title, WorkItemType, State)";
}

// If all inputs are valid , retrieves the work item data from the project
// Otherwise, throws an error.
static async Task<string> getAzureDevOpsData(string personalAccessToken, string getRequestUrl) {
    using (HttpClient client = new HttpClient())
    {
         client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
            Convert.ToBase64String(
                System.Text.ASCIIEncoding.ASCII.GetBytes(
                    string.Format("{0}:{1}", "", personalAccessToken))));
        using (HttpResponseMessage response = client.GetAsync(
            getRequestUrl
        ).Result)
        {
            response.EnsureSuccessStatusCode();
            string responseBody = await response.Content.ReadAsStringAsync();
            return responseBody;
        }
    }
}

// Adds all data to the lookup tables, and adds the header for the excelData
static void populateDefaultVariables(ref Dictionary<String, List<String>> workItemsData, ref Dictionary<String, String> workItemTypeLookup, ref Dictionary<String, bool> workItemAdded, ref String[][] excelData, ref string rawDataFromProject) {
    JsonDocument jsonDocument = JsonDocument.Parse(rawDataFromProject);
    JsonElement valueArray = jsonDocument.RootElement.GetProperty(JSON_DATA_VALUES);
    string extractedValue = valueArray.ToString();
    List<WorkItem> workItems = System.Text.Json.JsonSerializer.Deserialize<List<WorkItem>>(extractedValue);
    excelData = [.. excelData, [USER_NEED_WORK_ITEM, DESIGN_INPUT_WORK_ITEM, SRS_WORK_ITEM, TASK_WORK_ITEM]];
    foreach (var workItem in workItems) {
        List<String> children = new List<string>();
        for (int i = 0; i < workItem.Children.Count; i++) {
            children.Add(workItem.Children[i].WorkItemId.ToString());
        }
        workItemsData.Add(workItem.WorkItemId.ToString(), children);
        workItemTypeLookup.Add(workItem.WorkItemId.ToString(), workItem.WorkItemType);
        workItemAdded.Add(workItem.WorkItemId.ToString(), false);
    }
} 

// Organizes the data in excelData into a cascading sheet where all parents are 
// All non main work item types are added in "Other Work Items"
static void organiseExcelSheetData(Dictionary<string, string> workItemTypeLookup, Dictionary<string, List<string>> workItemsData, ref Dictionary<string, bool> workItemAdded, ref List<string> workItemIdsToAdd, ref string[][] excelData) {
    string[] workItemTypes = [USER_NEED_WORK_ITEM, DESIGN_INPUT_WORK_ITEM, SRS_WORK_ITEM, TASK_WORK_ITEM];
    foreach (string workItemType in workItemTypes) {
        foreach (KeyValuePair<String, String> ele in workItemTypeLookup) {
            if (ele.Value == workItemType && workItemAdded[ele.Key] == false) {
                appendItems(workItemsData, ref workItemAdded, workItemIdsToAdd, ref excelData, ele.Key);
            }
        }
        workItemIdsToAdd.Add("");          
    }

    excelData = [.. excelData, ["Other Work Items (with no parent object):"]];

    foreach (KeyValuePair<string, bool> workItem in workItemAdded) {
        if (workItem.Value == false) {
            string[] row = [workItem.Key, workItemTypeLookup[workItem.Key]];
            excelData = [.. excelData, row]; 
        }
    }
}

static void writeCsv(string filePath, params string[][] rows)
{
    using (StreamWriter writer = new StreamWriter(filePath))
    {
        for (int i = 0; i < rows.Length; i++)
        {
            writer.WriteLine(string.Join(",", rows[i]));
        }
    }
}

// Takes in data, and using the lookup table it stores the values in a table (2d array) foreach initial item in excelData
static void appendItems(Dictionary<String, List<String>> workItemsData, ref Dictionary<String, bool> workItemAdded, List<String> workItemIdsToAdd,ref String[][] excelData, String idToCheck) {
    List<String> children = [];
    workItemIdsToAdd.Add(idToCheck);
    workItemAdded[idToCheck] = true;
    String newworkitemids = "";
    for (int i = 0; i < workItemIdsToAdd.Count; i++) {
        newworkitemids = newworkitemids + workItemIdsToAdd[i] + ",";
    };

    if (workItemsData.TryGetValue(idToCheck, out children)) {
        // If we are at the leaf of the tree(no more children), this means that this is one row/"branch" in the excel sheet we need to store.            
        if (children.Any() == false) {
            string[] currentRow = [];
            //Append all data from the workItemsIdsToAdd to a String array
            foreach (var i in workItemIdsToAdd) {
                currentRow = [.. currentRow, i];
            }
            // Append current branch to the excel sheet data
            excelData = [.. excelData, currentRow];       
            //Pop the current node and continue searching other branches (if any)
            workItemIdsToAdd.RemoveAt(workItemIdsToAdd.Count - 1);
        }
        // Not at a leaf node, continue deeper into the current branch
        else {
            foreach (var child in children) {
                appendItems(workItemsData, ref workItemAdded, workItemIdsToAdd, ref excelData, child);
            }
            workItemIdsToAdd.RemoveAt(workItemIdsToAdd.Count - 1);
        }
    } else {
        Console.WriteLine("Item" + idToCheck + " does not exist in workItemsData");
    }
}

static async Task Execute()
{
    var apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
    var client = new SendGridClient(apiKey);
    var from =new EmailAddress("ganesh_ramasamy@dxdhub.sg", "This is Ganesh but somehow the intern is sending an email");
    var subject = "Sending with SendGrid is Fun";
    var to = new EmailAddress("ganesh_ramasamy@dxdhub.sg", "This is also Ganesh");
    var plainTextContent = "and easy to do anywhere, even with C#";
    var htmlContent = "<strong>and easy to do anywhere, even with C#</strong>";
    var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);
    var response = await client.SendEmailAsync(msg);
    Console.WriteLine("apikey is" + apiKey);
    Console.WriteLine("client is" + client);
    Console.WriteLine("from is" + from);
    Console.WriteLine("subject is" + subject);
    Console.WriteLine("to is" + to);
    Console.WriteLine("plainTextContent is" + plainTextContent);
    Console.WriteLine("htmlContent is" + htmlContent);
    Console.WriteLine("msg is" + msg);
    Console.WriteLine("response is" + response);
    Console.WriteLine("response status is" + response.StatusCode);
}

public class WorkItem
{
    public int WorkItemId { get; set; }
    public string Title { get; set; }
    public string WorkItemType { get; set; }
    public string State { get; set; }
    public List<WorkItem> Children { get; set; }
}



