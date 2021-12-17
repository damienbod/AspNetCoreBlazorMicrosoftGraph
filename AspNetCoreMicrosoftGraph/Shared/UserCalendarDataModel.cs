using System;
using System.ComponentModel.DataAnnotations;

namespace AspNetCoreMicrosoftGraph
{

    public class UserCalendarDataModel
    {
        [Required]
        public string Email { get; set; }
    }
}
