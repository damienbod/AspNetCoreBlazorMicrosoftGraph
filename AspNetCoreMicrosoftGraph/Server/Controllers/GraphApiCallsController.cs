using System;
using System.Collections.Generic;
using System.Linq;
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

        [HttpPost("MailboxSettings")]
        public async Task<IActionResult> MailboxSettings([FromBody] string email)
        {
            if (string.IsNullOrEmpty(email))
                return BadRequest("No email");
            try
            {
                var mailbox = await _graphApiClientService.GetUserMailboxSettings(email);

                if(mailbox == null)
                {
                    return NotFound($"mailbox settings for {email} not found");
                }
                var result = new List<MailboxSettingsData> {
                new MailboxSettingsData { Name = "User Email", Data = email },
                new MailboxSettingsData { Name = "AutomaticRepliesSetting", Data = mailbox.AutomaticRepliesSetting.Status.ToString() },
                new MailboxSettingsData { Name = "TimeZone", Data = mailbox.TimeZone },
                new MailboxSettingsData { Name = "Language", Data = mailbox.Language.DisplayName }
            };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPost("TeamsPresence")]
        public async Task<IActionResult> PresencePost([FromBody] string email)
        {
            if (string.IsNullOrEmpty(email))
                return BadRequest("No email");
            try
            {
                var userPresence = await _graphApiClientService.GetPresenceforEmail(email);

                if (userPresence.Count == 0)
                {
                    return NotFound(email);
                }

                var result = new List<PresenceData> {
                new PresenceData { Name = "User Email", Data = email },
                new PresenceData { Name = "Availability", Data = userPresence[0].Availability }
            };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("UserCalendar")]
        public async Task<IEnumerable<FilteredEventDto>> UserCalendar(UserCalendarDataModel userCalendarDataModel)
        {
            var userCalendar = await _graphApiClientService.GetCalanderForUser(
                User.Identity.Name,
                "2021-12-13T12:00:00-01:00",
                "2023-12-13T12:00:00-01:00");

            return userCalendar.Select(l => new FilteredEventDto
            {
                IsAllDay = l.IsAllDay.GetValueOrDefault(),
                Sensitivity = l.Sensitivity.ToString(),
                Start = l.Start?.DateTime,
                End = l.End?.DateTime,
                ShowAs = l.ShowAs.Value.ToString(),
                Subject=l.Subject
            });
        }
    }
}
