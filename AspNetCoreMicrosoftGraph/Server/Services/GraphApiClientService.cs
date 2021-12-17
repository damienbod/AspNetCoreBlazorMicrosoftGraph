using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AspNetCoreMicrosoftGraph.Server.Services
{
    public class GraphApiClientService
    {
        private readonly GraphServiceClient _graphServiceClient;

        public GraphApiClientService(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        public async Task<User> GetGraphApiUser(string email)
        {
            var upn = await GetUserIdAsync(email);
            if (!string.IsNullOrEmpty(upn))
                return null;

            return await _graphServiceClient.Users[upn]
                .Request()
                .GetAsync();
        }

        public async Task<MailboxSettings> GetUserMailboxSettings(string email)
        {
            var upn = await GetUserIdAsync(email);
            if (string.IsNullOrEmpty(upn))
                return null;

            var user = await _graphServiceClient.Users[upn]
                .Request()
                .Select("MailboxSettings")
                .GetAsync();

            return user.MailboxSettings;
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
            var upn = await GetUserIdAsync(email);
            if (string.IsNullOrEmpty(upn))
                return null;

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startDateTime", from),
                new QueryOption("endDateTime", to)
            };

            var calendarView = await _graphServiceClient.Users[upn].CalendarView
                .Request(queryOptions)
                .Select("start,end,subject,location,sensitivity, showAs, isAllDay")
                .GetAsync();

            return calendarView;
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
            var upn = await GetUserIdAsync(email);

            var ids = new List<string>()
            {
                upn
            };

            return await _graphServiceClient.Communications
                .GetPresencesByUserId(ids)
                .Request()
                .PostAsync();
        }
    }
}

