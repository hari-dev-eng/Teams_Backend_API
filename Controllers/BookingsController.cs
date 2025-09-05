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
    [HttpPost]
    //[Authorize]
    [Route("api/bookings")]
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
            var istZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
            var startIst = TimeZoneInfo.ConvertTimeFromUtc(dto.StartTime, istZone);
            var endIst = TimeZoneInfo.ConvertTimeFromUtc(dto.EndTime, istZone);

            // Handle all-day events
            DateTimeTimeZone startTimeZone;
            DateTimeTimeZone endTimeZone;

            if (dto.IsAllDay)
            {
                startTimeZone = new DateTimeTimeZone
                {
                    DateTime = startIst.ToString("yyyy-MM-ddT00:00:00"),
                    TimeZone = OutlookTz
                };
                endTimeZone = new DateTimeTimeZone
                {
                    DateTime = endIst.ToString("yyyy-MM-ddT23:59:59"),
                    TimeZone = OutlookTz
                };
            }
            else
            {
                startTimeZone = new DateTimeTimeZone
                {
                    DateTime = startIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = OutlookTz
                };
                endTimeZone = new DateTimeTimeZone
                {
                    DateTime = endIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = OutlookTz
                };
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
                ShowAs = showAsStatus,
                IsReminderOn = isReminderOn,
                ReminderMinutesBeforeStart = reminderMinutesBeforeStart,
                IsAllDay = dto.IsAllDay
            };

            // Handle recurrence if specified

            if (dto.IsRecurring && dto.RecurrencePattern != null)

            {

                @event.Recurrence = CreateRecurrencePattern(dto.RecurrencePattern, dto.StartTime);

            }

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
                isAllDay = createdEvent.IsAllDay,
                recurrence = createdEvent.Recurrence
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
    private PatternedRecurrence CreateRecurrencePattern(RecurrencePatternDto patternDto, DateTime startDate)
    {
        var recurrence = new PatternedRecurrence();

        // Pattern
        switch (patternDto.PatternType?.ToLower())
        {
            case "daily":
                recurrence.Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.Daily,
                    Interval = patternDto.Interval ?? 1
                };
                break;

            case "weekly":
                recurrence.Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.Weekly,
                    Interval = patternDto.Interval ?? 1,
                    DaysOfWeek = ConvertToGraphDaysOfWeek(patternDto.DaysOfWeek)
                };
                break;

            case "monthly":
                recurrence.Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.AbsoluteMonthly,
                    Interval = patternDto.Interval ?? 1,
                    DayOfMonth = patternDto.DayOfMonth ?? 1
                };
                break;

            case "yearly":
                recurrence.Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.AbsoluteYearly,
                    Interval = patternDto.Interval ?? 1,
                    Month = patternDto.Month ?? 1,
                    DayOfMonth = patternDto.DayOfMonth ?? 1
                };
                break;

            default:
                recurrence.Pattern = new RecurrencePattern
                {
                    Type = RecurrencePatternType.Daily,
                    Interval = 1
                };
                break;
        }

        // Range
        recurrence.Range = new RecurrenceRange
        {
            Type = patternDto.Range?.Type?.ToLower() switch
            {
                "enddate" => RecurrenceRangeType.EndDate,
                "numbered" => RecurrenceRangeType.Numbered,
                _ => RecurrenceRangeType.NoEnd
            },
            StartDate = new Date(startDate.Year, startDate.Month, startDate.Day),
            EndDate = patternDto.Range?.EndDate.HasValue == true
                ? new Date(patternDto.Range.EndDate.Value.Year,
                           patternDto.Range.EndDate.Value.Month,
                           patternDto.Range.EndDate.Value.Day)
                : null,
            NumberOfOccurrences = patternDto.Range?.NumberOfOccurrences
        };

        return recurrence;
    }

    private List<DayOfWeekObject?> ConvertToGraphDaysOfWeek(int daysOfWeekBitmask)
    {
        var days = new List<DayOfWeekObject?>();

        if ((daysOfWeekBitmask & 1) != 0) days.Add(DayOfWeekObject.Sunday);
        if ((daysOfWeekBitmask & 2) != 0) days.Add(DayOfWeekObject.Monday);
        if ((daysOfWeekBitmask & 4) != 0) days.Add(DayOfWeekObject.Tuesday);
        if ((daysOfWeekBitmask & 8) != 0) days.Add(DayOfWeekObject.Wednesday);
        if ((daysOfWeekBitmask & 16) != 0) days.Add(DayOfWeekObject.Thursday);
        if ((daysOfWeekBitmask & 32) != 0) days.Add(DayOfWeekObject.Friday);
        if ((daysOfWeekBitmask & 64) != 0) days.Add(DayOfWeekObject.Saturday);

        return days;
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
