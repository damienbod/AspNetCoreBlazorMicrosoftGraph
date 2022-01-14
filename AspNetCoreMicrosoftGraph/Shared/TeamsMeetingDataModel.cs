using System;
using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph
{
    public class TeamsMeetingDataModel
    {
        [Required]
        public string Attendees { get; set; }
        [Required]
        public string MeetingName { get; set; }

        [Required]
        public DateTime? From { get; set; }

        [Required]
        public DateTime? To { get; set; }
    }
}
