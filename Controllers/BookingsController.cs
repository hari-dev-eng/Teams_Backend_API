using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Azure.Identity;
using Teams_Backend_API.Models.DTOs;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using Azure.Core;
using Microsoft.Kiota.Abstractions;

namespace Teams_Backend_API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class BookingsController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _config;

    // you had these fields already; we will set them from the DTO before building the event
    private FreeBusyStatus? showAsStatus;
    private bool? isReminderOn;
    private int? reminderMinutesBeforeStart;

    public BookingsController(IConfiguration config)
    {
        _config = config;
        var tenantId = config["AzureAd:TenantId"];
        var clientId = config["AzureAd:ClientId"];
        var clientSecret = config["AzureAd:ClientSecret"];

        var clientSecretCredential = new ClientSecretCredential(
            tenantId,
            clientId,
            clientSecret
        );

        _graphClient = new GraphServiceClient(clientSecretCredential);
    }

    // ===========================
    // ADDED: helpers (private)
    // ===========================

    // Map "Busy"/"Free"/"Tentative" -> Graph status
    private static FreeBusyStatus MapShowAs(string? category) =>
     (category ?? "Busy").ToLowerInvariant() switch
     {
         "free" => FreeBusyStatus.Free,
         "tentative" => FreeBusyStatus.Tentative,
         _ => FreeBusyStatus.Busy
     };

    // Bitmask -> Graph DayOfWeek list
    // Bit positions: 0=Sunday,1=Monday,...,6=Saturday (matches Teams UX)
    private static List<DayOfWeekObject?> DaysOfWeekFromMask(int mask)
    {
        var list = new List<DayOfWeekObject?>();
        if ((mask & (1 << 0)) != 0) list.Add(DayOfWeekObject.Sunday);
        if ((mask & (1 << 1)) != 0) list.Add(DayOfWeekObject.Monday);
        if ((mask & (1 << 2)) != 0) list.Add(DayOfWeekObject.Tuesday);
        if ((mask & (1 << 3)) != 0) list.Add(DayOfWeekObject.Wednesday);
        if ((mask & (1 << 4)) != 0) list.Add(DayOfWeekObject.Thursday);
        if ((mask & (1 << 5)) != 0) list.Add(DayOfWeekObject.Friday);
        if ((mask & (1 << 6)) != 0) list.Add(DayOfWeekObject.Saturday);
        return list;
    }


    // Your DTO -> Graph PatternedRecurrence
    private static PatternedRecurrence? BuildGraphRecurrence(RecurrencePatternDto? src, DateOnly seriesStartLocal)
    {
        if (src == null) return null;

        var pattern = new RecurrencePattern
        {
            Interval = Math.Max(1, src.Interval ?? 1)
        };

        switch ((src.PatternType ?? "").ToLowerInvariant())
        {
            case "daily":
                pattern.Type = RecurrencePatternType.Daily;
                break;

            case "weekly":
                pattern.Type = RecurrencePatternType.Weekly;
                pattern.DaysOfWeek = DaysOfWeekFromMask(src.DaysOfWeek);
                break;

            case "monthly": // absolute day-of-month
                pattern.Type = RecurrencePatternType.AbsoluteMonthly;
                pattern.DayOfMonth = src.DayOfMonth ?? seriesStartLocal.Day;
                break;

            case "yearly": // absolute month/day
                pattern.Type = RecurrencePatternType.AbsoluteYearly;
                pattern.Month = src.Month ?? seriesStartLocal.Month; // 1..12
                pattern.DayOfMonth = src.DayOfMonth ?? seriesStartLocal.Day;
                break;

            default:
                return null;
        }

        var range = new RecurrenceRange
        {
            // Microsoft.Graph Date(year, month, day)
            StartDate = new Date(seriesStartLocal.Year, seriesStartLocal.Month, seriesStartLocal.Day)
        };

        var rangeType = (src.Range?.Type ?? "noEnd").ToLowerInvariant();
        if (rangeType is "enddate" or "date")
        {
            range.Type = RecurrenceRangeType.EndDate;
            if (src.Range!.EndDate.HasValue)
            {
                var ed = src.Range.EndDate.Value;
                range.EndDate = new Date(ed.Year, ed.Month, ed.Day);
            }
        }
        else if (rangeType is "numbered" or "after")
        {
            range.Type = RecurrenceRangeType.Numbered;
            range.NumberOfOccurrences = Math.Max(1, src.Range!.NumberOfOccurrences ?? 1);
        }
        else
        {
            range.Type = RecurrenceRangeType.NoEnd;
        }

        return new PatternedRecurrence { Pattern = pattern, Range = range };
    }
    // ===========================

    [HttpPost]
    //[Authorize]
    public async Task<IActionResult> Create([FromBody] BookingDto dto)
    {
        try
        {
            if (dto == null) return BadRequest("Request body is null || No data sent to API");

            var userEmail = dto.UserEmail;
            var roomEmail = dto.RoomEmail;

            if (string.IsNullOrEmpty(userEmail) || string.IsNullOrEmpty(roomEmail))
            {
                return BadRequest(new { error = "UserEmail and RoomEmail are required" });
            }

            // Collect attendees (excluding organizer & room)
            var attendees = new List<Attendee>();

            // Organizer
            attendees.Add(new Attendee
            {
                EmailAddress = new EmailAddress
                {
                    Address = userEmail,
                    Name = string.IsNullOrEmpty(dto.OrganizerName) ? userEmail.Split('@')[0] : dto.OrganizerName
                },
                Type = AttendeeType.Required
            });

            // Room as resource attendee
            attendees.Add(new Attendee
            {
                EmailAddress = new EmailAddress
                {
                    Address = roomEmail,
                    Name = dto.RoomName
                },
                Type = AttendeeType.Resource
            });

            // Extra attendees
            if (dto.Attendees != null && dto.Attendees.Any())
            {
                foreach (var att in dto.Attendees)
                {
                    if (!string.IsNullOrEmpty(att.Email))
                    {
                        attendees.Add(new Attendee
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = att.Email,
                                Name = string.IsNullOrEmpty(att.Name) ? att.Email.Split('@')[0] : att.Name
                            },
                            Type = AttendeeType.Required
                        });
                    }
                }
            }

            const string OutlookTz = "India Standard Time";

            // Convert input time (assume dto.StartTime and dto.EndTime are UTC)
            var istZone = TimeZoneInfo.FindSystemTimeZoneById(OutlookTz);
            var startIst = TimeZoneInfo.ConvertTimeFromUtc(dto.StartTime, istZone);
            var endIst = TimeZoneInfo.ConvertTimeFromUtc(dto.EndTime, istZone);

            // ===========================
            // ADDED: category & reminder mapping
            // ===========================
            showAsStatus = MapShowAs(dto.Category);
            isReminderOn = dto.Reminder > 0;
            reminderMinutesBeforeStart = dto.Reminder > 0 ? dto.Reminder : 0;

            // ===========================
            // ADDED: recurrence (only if requested)
            // ===========================
            PatternedRecurrence? recurrence = null;
            if (dto.IsRecurring && dto.RecurrencePattern != null)
            {
                // Graph wants the local calendar date of the first occurrence
                var startDateLocal = DateOnly.FromDateTime(startIst.Date);
                recurrence = BuildGraphRecurrence(dto.RecurrencePattern, startDateLocal);
            }

            var @event = new Event
            {
                Subject = dto.Title,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = dto.Description ?? dto.Title
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = startIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = OutlookTz
                },
                End = new DateTimeTimeZone
                {
                    DateTime = endIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = OutlookTz
                },
                Location = new Location
                {
                    DisplayName = dto.RoomName,
                    LocationEmailAddress = roomEmail
                },
                Attendees = attendees,
                IsOnlineMeeting = false,
                AllowNewTimeProposals = false,

                // use fields we set above to avoid touching your existing property names
                ShowAs = showAsStatus,
                IsReminderOn = isReminderOn,
                ReminderMinutesBeforeStart = reminderMinutesBeforeStart,

                // ADDED: recurrence
                Recurrence = recurrence,

                // (optional but harmless if you already normalize times for all-day on FE)
                IsAllDay = dto.IsAllDay
            };

            // Organizer = userEmail
            var createdEvent = await _graphClient.Users[userEmail]
                .Calendar
                .Events
                .PostAsync(@event);

            return Ok(new
            {
                id = createdEvent.Id,
                subject = createdEvent.Subject,
                start = createdEvent.Start,
                end = createdEvent.End,
                location = createdEvent.Location?.DisplayName,
                locationEmail = createdEvent.Location?.LocationEmailAddress,
                showAs = createdEvent.ShowAs,
                isReminderOn = createdEvent.IsReminderOn,
                reminderMinutesBeforeStart = createdEvent.ReminderMinutesBeforeStart,
                recurrence = createdEvent.Recurrence // echo back for debugging
            });
        }
        catch (ODataError ex)
        {
            var errorDetails = ex.Error?.Message ?? "Unknown Graph API error";
            if (ex.ResponseStatusCode == 409)
            {
                return Conflict(new { error = "Time conflict: The room is already booked at this time" });
            }
            return BadRequest(new { error = errorDetails });
        }
    }

    [HttpGet("GetAccessToken")]
    public async Task<IActionResult> GetAccessToken()
    {
        var clientId = _config["AzureAd:ClientId"];
        var tenantId = _config["AzureAd:TenantId"];
        var clientSecret = _config["AzureAd:ClientSecret"];
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        var app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .Build();

        var token = await app.AcquireTokenForClient(scopes).ExecuteAsync();
        return Ok(new { access_token = token.AccessToken });
    }
}
