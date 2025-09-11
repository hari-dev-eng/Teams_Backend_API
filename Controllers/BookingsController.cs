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
using System.Linq;

using GraphDayOfWeek = Microsoft.Graph.Models.DayOfWeekObject;
using KiotaDate = Microsoft.Kiota.Abstractions.Date;

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

    private static FreeBusyStatus MapShowAs(string? category) =>
     (category ?? "Busy").ToLowerInvariant() switch
     {
         "free" => FreeBusyStatus.Free,
         "tentative" => FreeBusyStatus.Tentative,
         _ => FreeBusyStatus.Busy
     };

    private static List<GraphDayOfWeek?> DaysOfWeekFromMask(int mask)
    {
        var list = new List<GraphDayOfWeek?>();
        if ((mask & (1 << 0)) != 0) list.Add(GraphDayOfWeek.Sunday);
        if ((mask & (1 << 1)) != 0) list.Add(GraphDayOfWeek.Monday);
        if ((mask & (1 << 2)) != 0) list.Add(GraphDayOfWeek.Tuesday);
        if ((mask & (1 << 3)) != 0) list.Add(GraphDayOfWeek.Wednesday);
        if ((mask & (1 << 4)) != 0) list.Add(GraphDayOfWeek.Thursday);
        if ((mask & (1 << 5)) != 0) list.Add(GraphDayOfWeek.Friday);
        if ((mask & (1 << 6)) != 0) list.Add(GraphDayOfWeek.Saturday);
        return list;
    }

    // Convert System.DayOfWeek -> Graph enum
    private static GraphDayOfWeek ToGraphDayOfWeek(System.DayOfWeek d) => d switch
    {
        System.DayOfWeek.Sunday => GraphDayOfWeek.Sunday,
        System.DayOfWeek.Monday => GraphDayOfWeek.Monday,
        System.DayOfWeek.Tuesday => GraphDayOfWeek.Tuesday,
        System.DayOfWeek.Wednesday => GraphDayOfWeek.Wednesday,
        System.DayOfWeek.Thursday => GraphDayOfWeek.Thursday,
        System.DayOfWeek.Friday => GraphDayOfWeek.Friday,
        System.DayOfWeek.Saturday => GraphDayOfWeek.Saturday,
        _ => GraphDayOfWeek.Monday
    };

    // Align the first occurrence date to the first selected weekday on/after baseLocal (weekly patterns)
    private static DateTime AlignStartToWeeklyMask(DateTime baseLocal, int mask)
    {
        if (mask == 0) return baseLocal; // no days selected -> no shift
        int baseDow = (int)baseLocal.DayOfWeek; // Sunday=0..Saturday=6
        for (int offset = 0; offset < 7; offset++)
        {
            int candidateDow = (baseDow + offset) % 7;
            if ((mask & (1 << candidateDow)) != 0)
            {
                return baseLocal.Date.AddDays(offset).Add(baseLocal.TimeOfDay);
            }
        }
        return baseLocal;
    }

    // DTO -> Graph PatternedRecurrence
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
                var dows = DaysOfWeekFromMask(src.DaysOfWeek);
                // Fallback: if no days selected, use the weekday from seriesStartLocal
                if (dows.Count == 0)
                {
                    var sysDow = new DateTime(seriesStartLocal.Year, seriesStartLocal.Month, seriesStartLocal.Day).DayOfWeek;
                    dows = new List<GraphDayOfWeek?> { ToGraphDayOfWeek(sysDow) };
                }
                pattern.DaysOfWeek = dows; // now List<...> matches property type
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
            StartDate = new KiotaDate(seriesStartLocal.Year, seriesStartLocal.Month, seriesStartLocal.Day)
        };

        var rangeType = (src.Range?.Type ?? "noEnd").ToLowerInvariant();
        if (rangeType is "enddate" or "date")
        {
            range.Type = RecurrenceRangeType.EndDate;
            if (src.Range!.EndDate.HasValue)
            {
                var ed = src.Range.EndDate.Value;
                range.EndDate = new KiotaDate(ed.Year, ed.Month, ed.Day);
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

            // category & reminder mapping
            showAsStatus = MapShowAs(dto.Category);
            isReminderOn = dto.Reminder > 0;
            reminderMinutesBeforeStart = dto.Reminder > 0 ? dto.Reminder : 0;

            // Align weekly so first occurrence lands on selected weekday(s)
            var alignedStartIst = startIst;
            var alignedEndIst = endIst;
            if (dto.IsRecurring && dto.RecurrencePattern != null &&
                string.Equals(dto.RecurrencePattern.PatternType, "weekly", StringComparison.OrdinalIgnoreCase))
            {
                var duration = endIst - startIst;
                alignedStartIst = AlignStartToWeeklyMask(startIst, dto.RecurrencePattern.DaysOfWeek);
                alignedEndIst = alignedStartIst + duration;
            }

            // Build recurrence (if any) using the aligned date
            PatternedRecurrence? recurrence = null;
            if (dto.IsRecurring && dto.RecurrencePattern != null)
            {
                var startDateLocal = DateOnly.FromDateTime(alignedStartIst.Date);
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
                    DateTime = alignedStartIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = OutlookTz
                },
                End = new DateTimeTimeZone
                {
                    DateTime = alignedEndIst.ToString("yyyy-MM-ddTHH:mm:ss"),
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

                ShowAs = showAsStatus,
                IsReminderOn = isReminderOn,
                ReminderMinutesBeforeStart = reminderMinutesBeforeStart,

                Recurrence = recurrence,
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
