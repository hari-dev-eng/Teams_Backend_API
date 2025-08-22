namespace TeamsMeetingViewer.Models
{
    public class MeetingViewModel
    {
        public string? Subject { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Organizer { get; set; }
        public string? OrganizerEmail { get; internal set; }
        public string? Location { get; internal set; }
    }
}