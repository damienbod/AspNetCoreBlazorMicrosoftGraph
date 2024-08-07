﻿using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

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
        var events = await GetCalanderForUserUsingGraph(email, from, to);

        var allEvents = new List<FilteredEvent>();

        foreach (var calenderEvent in events!)
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

        return allEvents;
    }

    private async Task<List<Event>?> GetCalanderForUserUsingGraph(string email, string from, string to)
    {
        var graphServiceClient = GetGraphClient();

        var id = await GetUserIdAsync(email, graphServiceClient);
        if (string.IsNullOrEmpty(id))
            return null;

        var calendarView = await graphServiceClient.Users[id].CalendarView
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new string[]
                { "start", "end", "subject", "location", "sensitivity", "showAs", "isAllDay" };
                requestConfiguration.QueryParameters.StartDateTime = from;
                requestConfiguration.QueryParameters.EndDateTime = to;
            });

        return calendarView!.Value;
    }

    public async Task<MailboxSettings?> GetUserMailboxSettings(string email)
    {
        var graphServiceClient = GetGraphClient();

        var id = await GetUserIdAsync(email, graphServiceClient);
        if (string.IsNullOrEmpty(id))
            return null;

        var userMailboxSettings = await graphServiceClient.Users[id]
            .MailboxSettings
            .GetAsync();

        return userMailboxSettings;
    }

    private static async Task<string> GetUserIdAsync(string email, GraphServiceClient graphServiceClient)
    {
        // Add a fix for external users
        var filter = $"userPrincipalName eq '{email}'";
        //var filter = $"startswith(userPrincipalName,'{email}')";

        var users = await graphServiceClient.Users.GetAsync((requestConfiguration) =>
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
