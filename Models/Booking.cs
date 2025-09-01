namespace Teams_Backend_API.Models
{
    public class Booking
    {
        public Guid Id { get; }
        public string Title { get; }
        public DateTime StartTime { get; }
        public DateTime EndTime { get; }
        public string UserEmail { get; }
        public string UserName { get; }
        public string Location { get; }
        public DateTime CreatedAt { get; }

        public Booking(Guid id, string title, DateTime startTime, DateTime endTime,
                      string userEmail, string userName, string location, DateTime createdAt)
        {
            Id = id;
            Title = title;
            StartTime = startTime;
            EndTime = endTime;
            UserEmail = userEmail;
            UserName = userName;
            Location = location;
            CreatedAt = createdAt;
        }
    }
}
