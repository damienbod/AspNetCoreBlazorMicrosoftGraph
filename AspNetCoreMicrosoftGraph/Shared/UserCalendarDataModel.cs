using System;
using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph
{

    public class UserCalendarDataModel
    {
        [Required]
        public string Email { get; set; }

        [Required]
        public string From { get; set; }

        [Required]
        public string To { get; set; }
    }
}
