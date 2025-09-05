using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace Teams_Backend_API.Models.DTOs
{
    public class BookingDto
    {
        public string ?UserEmail { get; set; }
        public string ?RoomEmail { get; set; }
        public string ?Title { get; set; }
        public string ?Description { get; set; }
        public string ?Location { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public List<AttendeeDto> ?Attendees { get; set; }

        [JsonPropertyName("category")]
        public string ?Category { get; set; }
        public string ?RoomName { get; set; }

        [JsonPropertyName("reminder")]
        public int Reminder { get; set; }
        public string? OrganizerName { get; internal set; }

        public bool IsAllDay { get; set; }
        public bool IsRecurring { get; set; }
        public RecurrencePatternDto ?RecurrencePattern { get; set; }
    }
    public class RecurrencePatternDto
    {
        public string ?PatternType { get; set; }
        public int? Interval { get; set; }
        public int? DayOfMonth { get; set; }
        public int? Month { get; set; }
        public int DaysOfWeek { get; set; } // bitmask
        public RecurrenceRangeDto ?Range { get; set; }
    }

    public class RecurrenceRangeDto
    {
        public string ?Type { get; set; }
        public DateTime? EndDate { get; set; }
        public int? NumberOfOccurrences { get; set; }
    }
}