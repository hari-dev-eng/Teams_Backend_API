using OutLook_Events;

namespace Teams_Backend_API.Models.Entities
{
    public class MeetingViewModel
    {
        public string? Subject { get; set; }
        public string? StartTime { get; set; }   // "yyyy-MM-ddTHH:mm:ss" (IST)
        public string? EndTime { get; set; }     // "yyyy-MM-ddTHH:mm:ss" (IST)
        public string? Organizer { get; set; }
        public string? OrganizerEmail { get; set; }
        public string? Location { get; set; }
        public int AttendeeCount { get; set; }
        public string? Id { get; internal set; }
        public string? ICalUId { get; internal set; }
    }
}