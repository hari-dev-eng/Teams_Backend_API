using System.Globalization;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System.Collections.Concurrent;
using Teams_Backend_API.Models.Entities;

namespace OutLook_Events
{

    [Route("api/[controller]")]
    [ApiController]
    public class MeetingsController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<MeetingsController> _logger;
        private readonly IHttpClientFactory _httpClientFactory;

        public MeetingsController(
            IConfiguration config,
            ILogger<MeetingsController> logger,
            IHttpClientFactory httpClientFactory)
        {
            _config = config;
            _logger = logger;
            _httpClientFactory = httpClientFactory;
        }
        //Create meeting
        [HttpGet]
        public async Task<IActionResult> GetMeetings(
            [FromQuery(Name = "userEmails")] string[]? userEmails,
            [FromQuery] string? date = null)
        {
            // Fallback if client sends userEmail instead of userEmails (backward compat)
            if ((userEmails == null || userEmails.Length == 0) && Request.Query.TryGetValue("userEmail", out var single))
            {
                userEmails = new[] { single.ToString() };
            }

            // Default emails if none are provided
            if (userEmails == null || userEmails.Length == 0)
            {
                userEmails = new[]
                {
                      "ffmeeting@conservesolution.com",
                      "gfmeeting@conservesolution.com",
                      "sfmeeting@conservesolution.com",
                      "contconference@conservesolution.com"

                };
            }

            if (userEmails.Any(e => !e.Contains("@")))
                return BadRequest(new { status = "failure", message = "Invalid email addresses." });

            // Validate Azure AD config (prevents “works local, fails in host” when env vars are missing)
            var clientId = _config["AzureAd:ClientId"];
            var clientSecret = _config["AzureAd:ClientSecret"];
            var tenantId = _config["AzureAd:TenantId"];
            if (string.IsNullOrWhiteSpace(clientId) ||
                string.IsNullOrWhiteSpace(clientSecret) ||
                string.IsNullOrWhiteSpace(tenantId))
            {
                _logger.LogError("Azure AD configuration missing (ClientId/ClientSecret/TenantId).");
                return StatusCode(500, new { status = "failure", message = "Azure AD configuration missing in server." });
            }

            try
            {
                // Acquire Graph token (client credentials)
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                    .Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                    .Build();

                var token = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                // Time handling — ASK GRAPH to return IST directly
                const string OutlookTz = "India Standard Time";
                var istZone = TimeZoneInfo.FindSystemTimeZoneById(OutlookTz);

                // Parse the requested date with multiple accepted formats
                var acceptedFormats = new[]
                {
                    "dd-M-yyyy", "d-M-yyyy", "dd-MM-yyyy",
                    "yyyy-MM-dd", "M/d/yyyy", "d/M/yyyy"
                };

                DateTime selectedDateIst;
                if (!string.IsNullOrWhiteSpace(date) &&
                    DateTime.TryParseExact(date, acceptedFormats, CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var parsed))
                {
                    // Treat parsed as IST calendar date (no time)
                    selectedDateIst = parsed.Date;
                }
                else
                {
                    selectedDateIst = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, istZone).Date;
                }

                // Build the start/end of day in IST (strings without zone; Graph will interpret as IST due to Prefer header)
                var startOfDayIst = new DateTime(selectedDateIst.Year, selectedDateIst.Month, selectedDateIst.Day, 0, 0, 0);
                var endOfDayIst = startOfDayIst.AddDays(1).AddSeconds(-1);

                string fmt(DateTime dt) => dt.ToString("yyyy-MM-dd'T'HH:mm:ss"); // do NOT append 'Z'

                var allowedLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Ground Floor Meeting Room",
                    "1st Floor Meeting Room",
                    "Conference Room",
                    "3rd Floor Meeting Room"
                };

                var allMeetings = new ConcurrentBag<MeetingViewModel>();

                // Query each mailbox in parallel using /calendarView (range-filtered) and IST preference
                var tasks = userEmails.Select(async email =>
                {
                    try
                    {
                        using var httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Authorization =
                            new AuthenticationHeaderValue("Bearer", token.AccessToken);

                        // Ask Graph to return times in IST so we don't perform any UTC conversion
                        httpClient.DefaultRequestHeaders.Add("Prefer", $"outlook.timezone=\"{OutlookTz}\"");

                        var endpoint =
                            $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(email)}/calendar/calendarView" +
                            $"?startDateTime={Uri.EscapeDataString(fmt(startOfDayIst))}" +
                            $"&endDateTime={Uri.EscapeDataString(fmt(endOfDayIst))}" +
                            $"&$top=200&$orderby=start/dateTime" +
                            "&$select=subject,organizer,start,end,location,attendees";

                        var response = await httpClient.GetAsync(endpoint);

                        if (!response.IsSuccessStatusCode)
                        {
                            _logger.LogWarning("Graph request failed for {Email} with {Code}: {Reason}",
                                email, (int)response.StatusCode, response.ReasonPhrase);
                            return;
                        }

                        var json = await response.Content.ReadAsStringAsync();
                        var parsedJson = JObject.Parse(json);
                        var events = parsedJson["value"];
                        if (events == null || !events.Any()) return;

                        foreach (var ev in events)
                        {
                            //_logger.LogInformation("Event data: {EventData}", ev.ToString());

                            var location = ev.SelectToken("location.displayName")?.ToString()?.Trim();
                            if (string.IsNullOrWhiteSpace(location) || !allowedLocations.Contains(location))
                                continue;

                            // Times are already in IST due to Prefer header
                            var startStr = ev.SelectToken("start.dateTime")?.ToString();
                            var endStr = ev.SelectToken("end.dateTime")?.ToString();

                            if (!DateTime.TryParse(startStr, out var startIst) ||
                                !DateTime.TryParse(endStr, out var endIst))
                                continue;

                            // Extra guard: keep only the selected day (in case of multi-day events)
                            if (startIst.Date != selectedDateIst) continue;

                            int attendeeCount = 0;
                            var attendeesToken = ev.SelectToken("attendees");
                            if (attendeesToken is JArray arr) attendeeCount = arr.Count;

                            // Use consistent SelectToken approach for all properties
                            var subjectStr = ev.SelectToken("subject")?.ToString();
                            if (string.IsNullOrWhiteSpace(subjectStr))
                                subjectStr = "[No Title]";

                            var organizerName = ev.SelectToken("organizer.emailAddress.name")?.ToString();
                            var organizerEmail = ev.SelectToken("organizer.emailAddress.address")?.ToString();

                            var eventId = ev.SelectToken("id")?.ToString();   // <-- add this
                            
                            allMeetings.Add(new MeetingViewModel
                            {
                                Id = eventId,
                                Subject = subjectStr,
                                StartTime = startIst.ToString("yyyy-MM-dd'T'HH:mm:ss"),
                                EndTime = endIst.ToString("yyyy-MM-dd'T'HH:mm:ss"),
                                Organizer = organizerName,
                                OrganizerEmail = organizerEmail,
                                Location = location,
                                AttendeeCount = attendeeCount
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error fetching calendar for {Email}", email);
                    }
                });

                await Task.WhenAll(tasks);

                var meetings = allMeetings
                    .OrderBy(m => m.StartTime, StringComparer.Ordinal)
                    .ToList();

                return Ok(new { status = "success", count = meetings.Count, meetings });
            }
            catch (MsalServiceException msalEx)
            {
                _logger.LogError(msalEx, "Azure AD token acquisition failed.");
                return StatusCode(500, new { status = "failure", message = "Azure AD token acquisition failed." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in GetMeetings");
                return StatusCode(500, new { status = "failure", message = "Internal server error", details = ex.Message });
            }
        }
        // Delete meeting
        [HttpDelete("{eventId}")]
        public async Task<IActionResult> DeleteMeeting(
            string eventId,
            [FromQuery] string calendarEmail,
            [FromQuery] string signedInUser)
        {
            if (string.IsNullOrWhiteSpace(eventId) ||
                string.IsNullOrWhiteSpace(calendarEmail) ||
                string.IsNullOrWhiteSpace(signedInUser))
            {
                return BadRequest(new
                {
                    status = "failure",
                    message = "eventId, calendarEmail, and signedInUser are required."
                });
            }

            try
            {
                // Enforce security: only allow the organizer (signed-in user must match organizer mailbox)
                if (!calendarEmail.Equals(signedInUser, StringComparison.OrdinalIgnoreCase))
                {
                    return Forbid(); // 403
                }

                // Acquire Graph token with app identity
                var clientId = _config["AzureAd:ClientId"];
                var clientSecret = _config["AzureAd:ClientSecret"];
                var tenantId = _config["AzureAd:TenantId"];
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                    .Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                    .Build();

                var token = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token.AccessToken);

                // Build delete endpoint (user = organizer mailbox, not room mailbox)
                var endpoint = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(signedInUser)}/events/{Uri.EscapeDataString(eventId)}";


                var response = await httpClient.DeleteAsync(endpoint);

                if (!response.IsSuccessStatusCode)
                {
                    var errorBody = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning(
                        "Failed to delete event {EventId} from {Calendar}: {Status} {Body}",
                        eventId, calendarEmail, response.StatusCode, errorBody);

                    return StatusCode((int)response.StatusCode, new
                    {
                        status = "failure",
                        message = "Graph deletion failed",
                        details = errorBody
                    });
                }

                return Ok(new { status = "success", message = "Meeting cancelled successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting meeting {EventId} from {Calendar}", eventId, calendarEmail);
                return StatusCode(500, new
                {
                    status = "failure",
                    message = "Internal server error",
                    details = ex.Message
                });
            }
        }
    }
}
