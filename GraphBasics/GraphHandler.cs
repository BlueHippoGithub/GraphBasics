using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using System.ComponentModel.Design;

namespace GraphBasics;

public class GraphHandler
{
    public GraphServiceClient GraphClient { get; private set; }

    public GraphHandler(string tenantId, string clientId, string clientSecret)
    {
        GraphClient = CreateGraphClient(tenantId, clientId, clientSecret);
    }
    public GraphServiceClient CreateGraphClient(string tenantId, string clientId, string clientSecret)
    {
        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        var clientSecretCredential = new ClientSecretCredential(
            tenantId, clientId, clientSecret, options);
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        return new GraphServiceClient(clientSecretCredential, scopes);
    }

    public async Task<User?> GetUser(string userPrincipalName)
    {
        return await GraphClient.Users[userPrincipalName].GetAsync();
    }

    public async Task<(IEnumerable<Site>?, IEnumerable<Site>?)> GetSharepointSites()
    {
        var sites = (await GraphClient.Sites.GetAllSites.GetAsync())?.Value;
        if(sites == null)
        {
            return (null, null);
        }

        sites.RemoveAll(x => string.IsNullOrEmpty(x.DisplayName));

        var spSites = new List<Site>();
        var oneDriveSites = new List<Site>();

        foreach (var site in sites)
        {
            if (site == null) continue;
            
            var compare = site.WebUrl?.Split(site.SiteCollection?.Hostname)[1].Split("/");
            if (compare.All(x => !string.IsNullOrEmpty(x)) || compare.Length < 1)
            {
                continue;
            }

            if (compare[1] == "sites" || string.IsNullOrEmpty(compare[1]))
                spSites.Add(site);
            else if (compare[1] == "personal")
                oneDriveSites.Add(site);
        }

        return (spSites, oneDriveSites);
    }
}
