namespace Teams_Backend_API.Models.DTOs
{
    public class MeetingResponse
    {
        public string EventId { get; set; } = "";
        public string Subject { get; set; } = "";
        public string StartTime { get; set; } = "";
        public string EndTime { get; set; } = "";
        public string? Organizer { get; set; }
        public string? OrganizerEmail { get; set; }
        public string Location { get; set; } = "";
        public int AttendeeCount { get; set; }
        public string CalendarEmail { get; set; } = "";
    }
}
