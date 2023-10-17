using GraphBasics;

string tenantId = "INPUT TENANT ID";
string clientId = "INPUT CLIENT ID";
string clientSecret = "INPUT CLIENT SECRET";

var graphHandler = new GraphHandler(tenantId, clientId, clientSecret);


Console.WriteLine("Get display name of user:");
var user = await graphHandler.GetUser("mail@example.com");
Console.WriteLine(user?.DisplayName);

Console.WriteLine("Get all sharepoint sites in tenant");
var spSites = (await graphHandler.GetSharepointSites()).Item1;
foreach (var site in spSites)
{
    Console.WriteLine(site.DisplayName);
}

Console.ReadLine();