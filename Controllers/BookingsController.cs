using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Azure.Identity;
using Teams_Backend_API.Models.DTOs;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

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
            
            // 2. Create the event with the room as an attendee
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
                Attendees = new List<Attendee>
                {
                    // The room as required attendee
                    new Attendee
                    {
                        EmailAddress = new EmailAddress {
                            Address = roomEmail,
                            Name = dto.Location
                        },
                        Type = AttendeeType.Required
                    },
                    // The user as required attendee
                    new Attendee
                    {
                       EmailAddress = new EmailAddress
                        {
                            Address = userEmail,
                            Name = userName
                        },
                        Type = AttendeeType.Required
                     }
                },
                // CRITICAL: Set these to false for physical room bookings
                IsOnlineMeeting = false,
                AllowNewTimeProposals = false
            };

            // 3. Create the event in the user's calendar
            var createdEvent = await _graphClient.Users[userEmail]
                .Calendar
                .Events
                .PostAsync(@event);

            // 4. Return the created event with location details
            return Ok(new
            {
                id = createdEvent.Id,
                subject = createdEvent.Subject,
                start = createdEvent.Start,
                end = createdEvent.End,
                location = createdEvent.Location?.DisplayName,
                locationEmail = createdEvent.Location?.LocationEmailAddress,
                onlineMeetingUrl = createdEvent.OnlineMeeting?.JoinUrl // Will be null for physical rooms
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

    // Helper method to find room email by display name
    private async Task<string?> FindRoomEmailByDisplayName(string roomDisplayName)
    {
        try
        {
            // Search for rooms (resource mailboxes)
            var rooms = await _graphClient.Users
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = $"userType eq 'Room' and displayName eq '{roomDisplayName}'";
                    requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "mail" };
                });

            var room = rooms?.Value?.FirstOrDefault();
            return room?.Mail;
        }
        catch
        {
            return null;
        }
    }

    // Additional method to get all available rooms
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

            return Ok(rooms?.Value?.Select(r => new {
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