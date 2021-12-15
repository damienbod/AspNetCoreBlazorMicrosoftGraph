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

            return await _graphServiceClient.Users[upn]
                .Request()
                .GetAsync();
        }

        public async Task<MailboxSettings> GetUserMailboxSettings(string email)
        {
            var upn = await GetUserIdAsync(email);

            var user = await _graphServiceClient.Users[upn]
                .Request()
                .Select("MailboxSettings")
                .GetAsync();

            return user.MailboxSettings;
        }
        

        private async Task<string> GetUserIdAsync(string email)
        {
            var filter = $"startswith(userPrincipalName,'{email}')";

            var users = await _graphServiceClient.Users
                .Request()
                .Filter(filter)
                .GetAsync();

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

