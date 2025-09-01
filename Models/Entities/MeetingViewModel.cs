using OutLook_Events;

namespace Teams_Backend_API.Models.Entities
{
    public class MeetingViewModel
    {
        public string? Subject { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Organizer { get; set; }
        public string? OrganizerEmail { get; set; }
        public string? Location { get; set; }

        public int AttendeeCount { get; set; }
    }
}