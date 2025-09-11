namespace Teams_Backend_API.Models.DTOs
{
    public class DeleteMeetingRequest
    {
        // Preferred path
        public string? EventId { get; set; }
        public string? CalendarEmail { get; set; }

        // Fallback composite
        public string? Subject { get; set; }
        public string? Organizer { get; set; }
        public DateTime StartTime { get; set; }
    }
}
