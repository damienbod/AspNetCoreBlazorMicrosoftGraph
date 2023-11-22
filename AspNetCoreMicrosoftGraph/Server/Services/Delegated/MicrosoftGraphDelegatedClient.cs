using Microsoft.Graph;
using Microsoft.Graph.Communications.GetPresencesByUserId;
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
        var presencesByUserIdPostResponse = await GetPresenceAsync(email);

        var allPresenceItems = new List<Presence>();

        if (presencesByUserIdPostResponse != null && presencesByUserIdPostResponse.Value!.Count > 0)
        {
            foreach (var presence in presencesByUserIdPostResponse.Value)
            {
                allPresenceItems.Add(presence);
            }
        }

        return allPresenceItems;
    }

    private async Task<GetPresencesByUserIdPostResponse?> GetPresenceAsync(string email)
    {
        var id = await GetUserIdAsync(email);

        var requestBody = new GetPresencesByUserIdPostRequestBody
        {
            Ids = new List<string>
            {
                id
            },
        };

        var result = await _graphServiceClient.Communications
            .GetPresencesByUserId
            .PostAsGetPresencesByUserIdPostResponseAsync(requestBody);

        return result;
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
