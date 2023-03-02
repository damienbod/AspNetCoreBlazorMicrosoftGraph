using AspNetCoreMicrosoftGraph.Server.Services;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;

namespace AspNetCoreMicrosoftGraph.Server.Controllers;

[ValidateAntiForgeryToken]
[Authorize(AuthenticationSchemes = CookieAuthenticationDefaults.AuthenticationScheme)]
[AuthorizeForScopes(Scopes = new string[] { "User.ReadBasic.All user.read" })]
[ApiController]
[Route("api/[controller]")]
public class GraphApiCallsController : ControllerBase
{
    private readonly MicrosoftGraphDelegatedClient _microsoftGraphDelegatedClient;
    private readonly MicrosoftGraphApplicationClient _microsoftGraphApplicationClient;
    private readonly TeamsService _teamsService;
    private readonly EmailService _emailService;

    public GraphApiCallsController(MicrosoftGraphDelegatedClient microsoftGraphDelegatedClient,
        MicrosoftGraphApplicationClient microsoftGraphApplicationClient,
        TeamsService teamsService,
        EmailService emailService)
    {
        _microsoftGraphDelegatedClient = microsoftGraphDelegatedClient;
        _microsoftGraphApplicationClient = microsoftGraphApplicationClient;
        _teamsService = teamsService;
        _emailService = emailService;
    }

    [HttpGet("UserProfile")]
    public async Task<IEnumerable<string>> UserProfile()
    {
        var userData = await _microsoftGraphDelegatedClient.GetGraphApiUser(User.Identity.Name);
        return new List<string> { $"DisplayName: {userData!.DisplayName}",
            $"GivenName: {userData.GivenName}", $"Preferred Language: {userData.PreferredLanguage}" };
    }

    [HttpPost("MailboxSettings")]
    public async Task<IActionResult> MailboxSettings([FromBody] string email)
    {
        if (string.IsNullOrEmpty(email))
            return BadRequest("No email");
        try
        {
            var mailbox = await _microsoftGraphApplicationClient.GetUserMailboxSettings(email);

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
            var userPresence = await _microsoftGraphDelegatedClient.GetPresenceforEmail(email);

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
        var userCalendar = await _microsoftGraphApplicationClient.GetCalanderForUser(
            userCalendarDataModel.Email, 
            userCalendarDataModel.From.Value.ToString("yyyy-MM-ddTHH:mm:ss.sssZ"),
            userCalendarDataModel.To.Value.ToString("yyyy-MM-ddTHH:mm:ss.sssZ"));

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

    [HttpPost("CreateTeamsMeeting")]
    public async Task<TeamsMeetingCreated> CreateTeamsMeeting(TeamsMeetingDataModel teamsMeetingDataModel)
    {
        var meeting = _teamsService.CreateTeamsMeeting(
            teamsMeetingDataModel.MeetingName,
            teamsMeetingDataModel.From.Value,
            teamsMeetingDataModel.To.Value);

        var attendees = teamsMeetingDataModel.Attendees.Split(';');
        List<string> items = new();
        items.AddRange(attendees);
        var updatedMeeting = _teamsService.AddMeetingParticipants(
          meeting, items);

        var createdMeeting = await _microsoftGraphDelegatedClient.CreateOnlineMeeting(updatedMeeting);

        var teamsMeetingCreated = new TeamsMeetingCreated
        {
            Subject = createdMeeting.Subject,
            JoinUrl = createdMeeting.JoinUrl,
            Attendees = createdMeeting.Participants.Attendees.Select(c => c.Upn).ToList()
        };

        // send emails
        foreach (var attendee in createdMeeting.Participants.Attendees)
        {
            var recipient = attendee.Upn.Trim();
            var message = _emailService.CreateStandardEmail(recipient,
                createdMeeting.Subject, createdMeeting.JoinUrl);

            await _microsoftGraphDelegatedClient.SendEmailAsync(message);
        }

        teamsMeetingCreated.EmailSent = "Emails sent to all attendees, please check your mailbox";

        return teamsMeetingCreated;
    }
    
}
