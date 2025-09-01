namespace Teams_Backend_API.Models.DTOs
{
    public class BookingDto
    {
        public string ?Title { get; set; }
        public string ?Description { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string ?Location { get; set; }
        public string ?UserEmail { get; set; }
        public string ?RoomEmail { get; set; }
        public List<AttendeeDto>? Attendees { get; set; }
    }

}