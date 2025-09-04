using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Azure.Identity;
using Teams_Backend_API.Models.DTOs;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using Azure.Core;

namespace Teams_Backend_API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class BookingsController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _config;

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

    public object Configuration { get; private set; }

    [HttpPost]
    //[Authorize]
    public async Task<IActionResult> Create([FromBody] BookingDto dto)
    {
        try
        {
            var userEmail = dto.UserEmail;
            var roomEmail = dto.RoomEmail;
            var userName = userEmail.Split('@')[0];

            if (string.IsNullOrEmpty(userEmail) || string.IsNullOrEmpty(roomEmail))
            {
                return BadRequest(new { error = "UserEmail and RoomEmail are required" });
            }
            if (dto == null) return BadRequest("Request body is null || No data sent to API");
            // Build attendees list dynamically
            var attendees = new List<Attendee>
            {
                // Room attendee
                new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = roomEmail,
                        Name = dto.Location
                    },
                    Type = AttendeeType.Required
                },
                // Organizer attendee
                new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = userEmail,
                        Name = userName
                    },
                    Type = AttendeeType.Required
                }
            };

            // Add extra attendees if provided
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

            //response options update
            if (!Enum.TryParse<FreeBusyStatus>(dto.category, out var showAsStatus))
            {
                // If parsing fails, default to 'Busy'
                showAsStatus = FreeBusyStatus.Oof;
            }

            // --- NEW: Handle the Reminder option directly ---
            // Assuming dto.Reminder is now an integer from the frontend (0 for None, >0 otherwise)
            int reminderMinutesBeforeStart = dto.reminder;

            // The IsReminderOn property should be true only if the reminder value is greater than zero.
            bool isReminderOn = dto.reminder > 0;

            // Create event
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
                    DateTime = dto.StartTime.ToString("o"),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = dto.EndTime.ToString("o"),
                    TimeZone = "UTC"
                },
                Location = new Location
                {
                    DisplayName = dto.Location,
                    LocationEmailAddress = roomEmail
                },
                Attendees = attendees,
                IsOnlineMeeting = false,
                AllowNewTimeProposals = false,

                ShowAs = showAsStatus,
                IsReminderOn = isReminderOn,
                ReminderMinutesBeforeStart = reminderMinutesBeforeStart
            };

            // Create the event in the user's calendar
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
                onlineMeetingUrl = createdEvent.OnlineMeeting?.JoinUrl,
                showAs = createdEvent.ShowAs,
                isReminderOn = createdEvent.IsReminderOn,
                reminderMinutesBeforeStart = createdEvent.ReminderMinutesBeforeStart
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
    [AllowAnonymous]
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

    [HttpGet("rooms")]
    [AllowAnonymous]
    public async Task<IActionResult> GetAvailableRooms()
    {
        try
        {
            var rooms = await _graphClient.Users
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "userType eq 'Room'";
                    requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "mail", "officeLocation" };
                });

            return Ok(rooms?.Value?.Select(r => new
            {
                id = r.Id,
                displayName = r.DisplayName,
                email = r.Mail,
                officeLocation = r.OfficeLocation
            }));
        }
        catch (ODataError ex)
        {
            return BadRequest(new { error = ex.Error?.Message });
        }
    }
}