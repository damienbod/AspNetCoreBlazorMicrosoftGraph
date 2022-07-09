using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph;

public class TeamsMeetingDataModel
{
    [Required]
    public string Attendees { get; set; } = string.Empty;
    [Required]
    public string MeetingName { get; set; } = string.Empty;

    [Required]
    public DateTime? From { get; set; }

    [Required]
    public DateTime? To { get; set; }
}
