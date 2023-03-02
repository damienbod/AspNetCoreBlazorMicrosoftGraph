using Microsoft.Graph;

namespace AspNetCoreMicrosoftGraph.Server.Services;

public class MicrosoftGraphDelegatedClient
{
    private readonly GraphServiceClient _graphServiceClient;

    public MicrosoftGraphDelegatedClient(GraphServiceClient graphServiceClient)
    {
        _graphServiceClient = graphServiceClient;
    }

    public async Task<User?> GetGraphApiUser(string? email)
    {
        if (email == null) return null;

        var id = await GetUserIdAsync(email);
        if (string.IsNullOrEmpty(id))
            return null;

        return await _graphServiceClient.Users[id]
            .Request()
            .GetAsync();
    }

    private async Task<string> GetUserIdAsync(string email)
    {
        var filter = $"userPrincipalName eq '{email}'";
        //var filter = $"startswith(userPrincipalName,'{email}')";

        var users = await _graphServiceClient.Users
            .Request()
            .Filter(filter)
            .GetAsync();

        if(users.CurrentPage.Count == 0)
        {
            return string.Empty;
        }
        return users.CurrentPage[0].Id;
    }

    public async Task<List<Presence>> GetPresenceforEmail(string email)
    {
        var cloudCommunicationPages = await GetPresenceAsync(email);

        var allPresenceItems = new List<Presence>();

        while (cloudCommunicationPages != null && cloudCommunicationPages.Count > 0)
        {
            foreach (var presence in cloudCommunicationPages)
            {
                allPresenceItems.Add(presence);
            }

            if (cloudCommunicationPages.NextPageRequest == null)
                break;
        }

        return allPresenceItems;
    }

    private async Task<ICloudCommunicationsGetPresencesByUserIdCollectionPage> GetPresenceAsync(string email)
    {
        var id = await GetUserIdAsync(email);

        var ids = new List<string>()
        {
            id
        };

        return await _graphServiceClient.Communications
            .GetPresencesByUserId(ids)
            .Request()
            .PostAsync();
    }

    public async Task SendEmailAsync(Message message)
    {
        var saveToSentItems = true;

        await _graphServiceClient.Me
            .SendMail(message, saveToSentItems)
            .Request()
            .PostAsync();
    }

    public async Task<OnlineMeeting> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings
            .Request()
            .AddAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeeting.Id]
            .Request()
            .UpdateAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting> GetOnlineMeeting(string onlineMeetingId)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeetingId]
            .Request()
            .GetAsync();
    }
}

