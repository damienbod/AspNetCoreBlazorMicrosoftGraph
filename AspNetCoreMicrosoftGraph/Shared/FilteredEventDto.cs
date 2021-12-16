using System;

namespace AspNetCoreMicrosoftGraph
{
    public class FilteredEventDto
    {
        public string Start { get; set; }
        public string End { get; set; }
        public string Subject { get; set; }
        public string Sensitivity { get; set; }
        public string ShowAs { get; set; }
        public bool IsAllDay { get; set; }  
    }
}
