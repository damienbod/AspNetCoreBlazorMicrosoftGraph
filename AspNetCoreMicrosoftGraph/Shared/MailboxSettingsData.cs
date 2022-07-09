using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph;

public class MailboxSettingsData
{
    public string? Name { get; set; }
    public string? Data { get; set; }
}

public class MailboxSettingsModel
{
    [Required]
    public string EmailMailboxSettings { get; set; } = string.Empty;
}
