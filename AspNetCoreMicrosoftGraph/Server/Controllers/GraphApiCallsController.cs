using System.Collections.Generic;
using System.Threading.Tasks;
using AspNetCoreMicrosoftGraph.Server.Services;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;

namespace AspNetCoreMicrosoftGraph.Server.Controllers
{
    [ValidateAntiForgeryToken]
    [Authorize(AuthenticationSchemes = CookieAuthenticationDefaults.AuthenticationScheme)]
    [AuthorizeForScopes(Scopes = new string[] { "User.ReadBasic.All user.read" })]
    [ApiController]
    [Route("api/[controller]")]
    public class GraphApiCallsController : ControllerBase
    {
        private GraphApiClientService _graphApiClientService;

        public GraphApiCallsController(GraphApiClientService graphApiClientService)
        {
            _graphApiClientService = graphApiClientService;
        }

        [HttpGet("UserProfile")]
        public async Task<IEnumerable<string>> UserProfile()
        {
            var userData = await _graphApiClientService.GetGraphApiUser(User.Identity.Name);
            return new List<string> { $"DisplayName: {userData.DisplayName}",
                $"GivenName: {userData.GivenName}", $"Preferred Language: {userData.PreferredLanguage}" };
        }

        [HttpGet("MailboxSettings")]
        public async Task<IEnumerable<string>> MailboxSettings()
        {
            var mailboxSettings = await _graphApiClientService.GetUserMailboxSettings(User.Identity.Name);
            return new List<string> { $"AutomaticRepliesSetting Status: {mailboxSettings.AutomaticRepliesSetting.Status}",
                $"TimeZone: {mailboxSettings.TimeZone}", $"Language: {mailboxSettings.Language.DisplayName}" };
        }

        [HttpGet("TeamsPresence")]
        public async Task<IEnumerable<string>> Presence()
        {
            var userPresence = await _graphApiClientService.GetPresenceforEmail(User.Identity.Name);
            return new List<string> { $"User Email: {User.Identity.Name}",
                $"Availability: {userPresence[0].Availability}" };
        }

        [HttpGet("UserCalendar")]
        public async Task<List<FilteredEvent>> UserCalendar()
        {
            var userCalendar = await _graphApiClientService.GetCalanderForUser(
                User.Identity.Name,
                "2021-12-13T12:00:00-01:00",
                "2023-12-13T12:00:00-01:00");

            return userCalendar;
        }
    }
}
