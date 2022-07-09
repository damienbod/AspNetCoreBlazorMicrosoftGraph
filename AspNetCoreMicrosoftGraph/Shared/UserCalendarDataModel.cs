using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph;

public class UserCalendarDataModel
{
    [Required]
    public string Email { get; set; } = string.Empty;

    [Required]
    public DateTime? From { get; set; }

    [Required]
    public DateTime? To { get; set; }
}
