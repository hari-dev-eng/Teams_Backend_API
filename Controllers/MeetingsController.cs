using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using TeamsMeetingViewer.Models;

namespace OutLook_Events
{
    [Route("api/[controller]")]
    [ApiController]
    public class MeetingsController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<MeetingsController> _logger;

        public MeetingsController(IConfiguration config, ILogger<MeetingsController> logger)
        {
            _config = config;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetMeetings([FromQuery] string userEmail)
        {
            var meetings = new List<MeetingViewModel>();

            // Validate input
            if (string.IsNullOrWhiteSpace(userEmail) || !userEmail.Contains("@"))
            {
                return BadRequest(new { status = "failure", message = "Invalid or missing email address." });
            }

            try
            {
                // Load secrets securely from appsettings.json / environment variables
                string clientId = _config["AzureAd:ClientId"] ?? throw new Exception("Missing ClientId");
                string clientSecret = _config["AzureAd:ClientSecret"] ?? throw new Exception("Missing ClientSecret");
                string tenantId = _config["AzureAd:TenantId"] ?? throw new Exception("Missing TenantId");
                string[] scopes = new[] { "https://graph.microsoft.com/.default" };

                // Authenticate
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                    .Build();

                AuthenticationResult result;
                try
                {
                    result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
                }
                catch (Exception authEx)
                {
                    _logger.LogError(authEx, "Failed to acquire Graph API token");
                    return StatusCode(401, new { status = "failure", message = "Unauthorized. Token acquisition failed." });
                }

                // Graph API call
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", result.AccessToken);

                string endpoint = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userEmail)}/events?$top=50&$orderby=start/dateTime";
                HttpResponseMessage response = await client.GetAsync(endpoint);

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogWarning("Graph API returned {StatusCode} for {Email}", response.StatusCode, userEmail);
                    return StatusCode((int)response.StatusCode, new
                    {
                        status = "failure",
                        message = "Graph API call failed",
                        details = await response.Content.ReadAsStringAsync()
                    });
                }

                string json = await response.Content.ReadAsStringAsync();
                JObject parsed;
                try
                {
                    parsed = JObject.Parse(json);
                }
                catch (Exception parseEx)
                {
                    _logger.LogError(parseEx, "Failed to parse Graph API response");
                    return StatusCode(500, new { status = "failure", message = "Invalid response format from Graph API" });
                }

                var events = parsed["value"];
                if (events == null || !events.Any())
                {
                    return Ok(new { status = "success", message = "No meetings found", meetings = new List<MeetingViewModel>() });
                }

                // Time zone conversion
                TimeZoneInfo istZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                var todayIst = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, istZone).Date;

                var allowedLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Ground Floor Meeting Room",
                    "1st Floor Meeting Room",
                    "Conference Room",
                    "3rd Floor Meeting Room"
                };

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

                    // THIS IS THE FIX: Only skip past meetings
                    if (startIst.Date < todayIst)
                        continue;

                    meetings.Add(new MeetingViewModel
                    {
                        Subject = ev["subject"]?.ToString(),
                        StartTime = startIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                        EndTime = endIst.ToString("yyyy-MM-ddTHH:mm:ss"),
                        Organizer = ev.SelectToken("organizer.emailAddress.name")?.ToString(),
                        OrganizerEmail = ev.SelectToken("organizer.emailAddress.address")?.ToString(),
                        Location = location
                    });
                }


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
