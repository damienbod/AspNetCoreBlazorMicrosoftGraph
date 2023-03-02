using Azure.Identity;
using Microsoft.Graph;

namespace AspNetCoreMicrosoftGraph.Server.Services;

public class MicrosoftGraphApplicationClient
{
    private readonly IConfiguration _configuration;

    public MicrosoftGraphApplicationClient(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public async Task<List<FilteredEvent>> GetCalanderForUser(string email, string from, string to)
    {
        var userCalendarViewCollectionPages = await GetCalanderForUserUsingGraph(email, from, to);

        var allEvents = new List<FilteredEvent>();

        while (userCalendarViewCollectionPages != null && userCalendarViewCollectionPages.Count > 0)
        {
            foreach (var calenderEvent in userCalendarViewCollectionPages)
            {
                var filteredEvent = new FilteredEvent
                {
                    ShowAs = calenderEvent.ShowAs,
                    Sensitivity = calenderEvent.Sensitivity,
                    Start = calenderEvent.Start,
                    End = calenderEvent.End,
                    Subject = calenderEvent.Subject,
                    IsAllDay = calenderEvent.IsAllDay,
                    Location = calenderEvent.Location
                };
                allEvents.Add(filteredEvent);
            }

            if (userCalendarViewCollectionPages.NextPageRequest == null)
                break;
        }

        return allEvents;
    }

    private async Task<IUserCalendarViewCollectionPage> GetCalanderForUserUsingGraph(string email, string from, string to)
    {
        var graphServiceClient = GetGraphClient();

        var id = await GetUserIdAsync(email, graphServiceClient);
        if (string.IsNullOrEmpty(id))
            return null;

        var queryOptions = new List<QueryOption>()
        {
            new QueryOption("startDateTime", from),
            new QueryOption("endDateTime", to)
        };

        var calendarView = await graphServiceClient.Users[id].CalendarView
            .Request(queryOptions)
            .Select("start,end,subject,location,sensitivity, showAs, isAllDay")
            .GetAsync();

        return calendarView;
    }

    public async Task<MailboxSettings> GetUserMailboxSettings(string email)
    {
        var graphServiceClient = GetGraphClient();

        var id = await GetUserIdAsync(email, graphServiceClient);
        if (string.IsNullOrEmpty(id))
            return null;

        var user = await graphServiceClient.Users[id]
            .Request()
            .Select("MailboxSettings")
            .GetAsync();

        return user.MailboxSettings;
    }

    private async Task<string> GetUserIdAsync(string email, GraphServiceClient graphServiceClient)
    {
        // Add a fix for external users
        var filter = $"userPrincipalName eq '{email}'";
        //var filter = $"startswith(userPrincipalName,'{email}')";

        var users = await graphServiceClient.Users
            .Request()
            .Filter(filter)
            .GetAsync();

        if (users.CurrentPage.Count == 0)
        {
            return string.Empty;
        }
        return users.CurrentPage[0].Id;
    }

    private GraphServiceClient GetGraphClient()
    {
        string[] scopes = new[] { "https://graph.microsoft.com/.default" };
        var tenantId = _configuration["AzureAd:TenantId"];

        // Values from app registration
        var clientId = _configuration.GetValue<string>("AzureAd:ClientId");
        var clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");

        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
        var clientSecretCredential = new ClientSecretCredential(
            tenantId, clientId, clientSecret, options);

        return new GraphServiceClient(clientSecretCredential, scopes);
    }

}
