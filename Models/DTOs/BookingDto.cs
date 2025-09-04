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

        [JsonPropertyName("reminder")]
        public int Reminder { get; set; }
    }

   
}