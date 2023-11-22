using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

namespace AspNetCoreMicrosoftGraph.Server.Services;

public class MicrosoftGraphDelegatedClient
{
    private readonly GraphServiceClient _graphServiceClient;

    public MicrosoftGraphDelegatedClient(GraphServiceClient graphServiceClient)
    {
        _graphServiceClient = graphServiceClient;
    }

    public async Task SendEmailAsync(Message message)
    {
        var saveToSentItems = true;

        var body = new SendMailPostRequestBody
        {
            Message = message,
            SaveToSentItems = saveToSentItems
        };

        await _graphServiceClient.Me.SendMail
            .PostAsync(body);
    }

    public async Task<OnlineMeeting?> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings
            .PostAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeeting.Id]
            .PatchAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> GetOnlineMeeting(string onlineMeetingId)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeetingId]
            .GetAsync();
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

    private async Task<string> GetUserIdAsync(string email)
    {
        // Add a fix for external users
        var filter = $"userPrincipalName eq '{email}'";
        //var filter = $"startswith(userPrincipalName,'{email}')";

        var users = await _graphServiceClient.Users.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Filter = filter;
        });

        var userId = users!.Value!.FirstOrDefault()!.Id;

        if (string.IsNullOrEmpty(userId))
        {
            return string.Empty;
        }

        return userId;
    }

    public async Task<User?> GetGraphApiUser(string? email)
    {
        if (email == null) return null;

        var id = await GetUserIdAsync(email);
        if (string.IsNullOrEmpty(id))
            return null;

        return await _graphServiceClient.Users[id]
            .GetAsync();
    }
}
