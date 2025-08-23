using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using TeamsMeetingViewer.Models;
using System.Collections.Concurrent;

namespace OutLook_Events
{
    [Route("api/[controller]")]
    [ApiController]
    public class MeetingsController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<MeetingsController> _logger;
        private readonly IHttpClientFactory _httpClientFactory;

        public MeetingsController(IConfiguration config, ILogger<MeetingsController> logger, IHttpClientFactory httpClientFactory)
        {
            _config = config;
            _logger = logger;
            _httpClientFactory = httpClientFactory;
        }

        [HttpGet]
        public async Task<IActionResult> GetMeetings([FromQuery] string[] userEmails, [FromQuery] string date = null)
        {
            var meetings = new List<MeetingViewModel>();

            // If no emails provided, use the default 4 emails
            if (userEmails == null || userEmails.Length == 0)
            {
                userEmails = new string[] {
                    "gfmeeting@conservesolution.com",
                    "ffmeeting@conservesolution.com",
                      "contconference@conservesolution.com",
                      "sfmeeting@conservesolution.com"
                };
            }

            if (userEmails.Any(e => !e.Contains("@")))
                return BadRequest(new { status = "failure", message = "Invalid email addresses." });

            try
            {
                string clientId = _config["AzureAd:ClientId"];
                string clientSecret = _config["AzureAd:ClientSecret"];
                string tenantId = _config["AzureAd:TenantId"];
                string[] scopes = new[] { "https://graph.microsoft.com/.default" };

                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                    .Build();

                var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                // Time zone setup
                TimeZoneInfo istZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");

                // Parse date from query OR fallback to today
                DateTime selectedDateIst;
                if (!string.IsNullOrEmpty(date) && DateTime.TryParse(date, out var parsedDate))
                    selectedDateIst = parsedDate.Date;
                else
                    selectedDateIst = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, istZone).Date;

                var allowedLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Ground Floor Meeting Room",
                    "1st Floor Meeting Room",
                    "Conference Room",
                    "3rd Floor Meeting Room"
                };

                // Fetch meetings from all emails in parallel
                var allMeetings = new ConcurrentBag<MeetingViewModel>();

                // Create a list of tasks for all email requests
                var tasks = userEmails.Select(async email =>
                {
                    try
                    {
                        // Create a new HttpClient for each request (better for parallel execution)
                        using var httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);

                        string endpoint = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(email)}/events?$top=50&$orderby=start/dateTime";
                        var response = await httpClient.GetAsync(endpoint);

                        if (!response.IsSuccessStatusCode)
                        {
                            _logger.LogWarning($"Failed to fetch events for {email}: {response.StatusCode}");
                            return;
                        }

                        var json = await response.Content.ReadAsStringAsync();
                        var parsed = JObject.Parse(json);
                        var events = parsed["value"];

                        if (events == null || !events.Any())
                            return;

                        foreach (var ev in events)
                        {
                            var location = ev.SelectToken("location.displayName")?.ToString()?.Trim();
                            if (string.IsNullOrWhiteSpace(location) || !allowedLocations.Contains(location))
                                continue;

                            if (!DateTime.TryParse(ev.SelectToken("start.dateTime")?.ToString(), out var startUtc) ||
                                !DateTime.TryParse(ev.SelectToken("end.dateTime")?.ToString(), out var endUtc))
                                continue;

                            DateTime startIst = TimeZoneInfo.ConvertTimeFromUtc(startUtc, istZone);
                            DateTime endIst = TimeZoneInfo.ConvertTimeFromUtc(endUtc, istZone);

                            // Only meetings for the selected date
                            if (startIst.Date != selectedDateIst)
                                continue;

                            // Get the attendees JSON array
                            var attendeesToken = ev.SelectToken("attendees");

                            // Calculate the attendee count
                            int attendeeCount = 0;
                            if (attendeesToken is JArray attendeesArray)
                            {
                                attendeeCount = attendeesArray.Count;
                            }

                            allMeetings.Add(new MeetingViewModel
                            {
                                Subject = ev["subject"]?.ToString(),
                                StartTime = startIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                                EndTime = endIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                                Organizer = ev.SelectToken("organizer.emailAddress.name")?.ToString(),
                                OrganizerEmail = ev.SelectToken("organizer.emailAddress.address")?.ToString(),
                                Location = location,
                                AttendeeCount = attendeeCount
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error processing events for {email}");
                    }
                });

                // Wait for all email requests to complete
                await Task.WhenAll(tasks);

                // Convert to list and order by start time
                meetings = allMeetings.OrderBy(m => m.StartTime).ToList();

                return Ok(new { status = "success", count = meetings.Count, meetings });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in GetMeetings");
                return StatusCode(500, new { status = "failure", message = "Internal server error", details = ex.Message });
            }
        }
    }
}