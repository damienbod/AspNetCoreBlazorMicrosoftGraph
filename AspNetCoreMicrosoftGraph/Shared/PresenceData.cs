using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph;

public class PresenceData
{
    public string? Name { get; set; }
    public string? Data { get; set; }
}

public class EmailPresenceModel
{
    [Required]
    public string EmailPresence { get; set; } = string.Empty;
}
