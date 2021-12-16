using Microsoft.Graph;

namespace AspNetCoreMicrosoftGraph.Server.Services
{
    public class FilteredEvent
    {
        public DateTimeTimeZone Start { get; set; }
        public DateTimeTimeZone End { get; set; }
        public string Subject { get; set; }
        public Location Location { get; set; }
        public Sensitivity? Sensitivity { get; set; }
        public FreeBusyStatus? ShowAs { get; set; }
        public bool? IsAllDay { get; set; }  

    }
}
