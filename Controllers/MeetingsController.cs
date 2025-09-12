using System.Globalization;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System.Collections.Concurrent;
using Teams_Backend_API.Models.Entities;
using System.Linq;
using System.Collections.Generic;

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

        [HttpGet]
        public async Task<IActionResult> GetMeetings(
            [FromQuery(Name = "userEmails")] string[]? userEmails,
            [FromQuery] string? date = null)
        {
            if ((userEmails == null || userEmails.Length == 0) && Request.Query.TryGetValue("userEmail", out var single))
            {
                userEmails = new[] { single.ToString() };
            }

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

            // Map each room mailbox to the friendly room name you show in the UI.
            var mailboxToRoomName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["gfmeeting@conservesolution.com"] = "Ground Floor Meeting Room",
                ["ffmeeting@conservesolution.com"] = "1st Floor Meeting Room",
                ["contconference@conservesolution.com"] = "Conference Room",
                ["sfmeeting@conservesolution.com"] = "3rd Floor Meeting Room",
            };

            var allowedLocations = new HashSet<string>(mailboxToRoomName.Values, StringComparer.OrdinalIgnoreCase);

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
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                    .Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                    .Build();

                var token = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                const string OutlookTz = "India Standard Time";
                var istZone = TimeZoneInfo.FindSystemTimeZoneById(OutlookTz);

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
                    selectedDateIst = parsed.Date;
                }
                else
                {
                    selectedDateIst = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, istZone).Date;
                }

                var startOfDayIst = new DateTime(selectedDateIst.Year, selectedDateIst.Month, selectedDateIst.Day, 0, 0, 0);
                var endOfDayIst = startOfDayIst.AddDays(1).AddSeconds(-1);
                string fmt(DateTime dt) => dt.ToString("yyyy-MM-dd'T'HH:mm:ss");

                var allMeetings = new ConcurrentBag<MeetingViewModel>();

                var tasks = userEmails.Select(async email =>
                {
                    try
                    {
                        using var httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Authorization =
                            new AuthenticationHeaderValue("Bearer", token.AccessToken);
                        httpClient.DefaultRequestHeaders.Add("Prefer", $"outlook.timezone=\"{OutlookTz}\"");

                        // NOTE: include 'locations' so we can see every room when multiple are selected
                        var endpoint =
                            $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(email)}/calendar/calendarView" +
                            $"?startDateTime={Uri.EscapeDataString(fmt(startOfDayIst))}" +
                            $"&endDateTime={Uri.EscapeDataString(fmt(endOfDayIst))}" +
                            $"&$top=200&$orderby=start/dateTime" +
                            "&$select=id,subject,organizer,start,end,location,locations,attendees,iCalUId";

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

                        // Friendly room for THIS mailbox
                        mailboxToRoomName.TryGetValue(email, out var thisMailboxRoomName);

                        foreach (var ev in events)
                        {
                            var startStr = ev.SelectToken("start.dateTime")?.ToString();
                            var endStr = ev.SelectToken("end.dateTime")?.ToString();

                            if (!DateTime.TryParse(startStr, out var startIst) ||
                                !DateTime.TryParse(endStr, out var endIst))
                                continue;

                            // keep the selected day (if you have cross-midnight events, relax this)
                            if (startIst.Date != selectedDateIst) continue;

                            // Collect every possible room name from Graph
                            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                            var primaryLocation = ev.SelectToken("location.displayName")?.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(primaryLocation))
                                names.Add(primaryLocation);

                            if (ev["locations"] is JArray jLocs)
                            {
                                foreach (var loc in jLocs)
                                {
                                    var ln = loc["displayName"]?.ToString()?.Trim();
                                    if (!string.IsNullOrWhiteSpace(ln)) names.Add(ln);
                                }
                            }

                            if (ev["attendees"] is JArray jAtt)
                            {
                                foreach (var a in jAtt)
                                {
                                    var type = a["type"]?.ToString();
                                    if (string.Equals(type, "resource", StringComparison.OrdinalIgnoreCase))
                                    {
                                        var rn = a["emailAddress"]?["name"]?.ToString()?.Trim();
                                        if (!string.IsNullOrWhiteSpace(rn)) names.Add(rn);
                                    }
                                }
                            }

                            // Intersect with our known rooms
                            var roomHits = names.Where(n => allowedLocations.Contains(n)).ToList();

                            // Decide if this event belongs to THIS mailbox' room
                            // - If we know the friendly name for this mailbox, include only when it’s in the event’s rooms
                            // - If Graph didn’t give names (edge), still include for this mailbox as a fallback
                            bool includeForThisMailbox =
                                (!string.IsNullOrEmpty(thisMailboxRoomName) && roomHits.Contains(thisMailboxRoomName))
                                || (string.IsNullOrEmpty(thisMailboxRoomName) && roomHits.Count > 0)
                                || (roomHits.Count == 0 && !string.IsNullOrEmpty(thisMailboxRoomName));

                            if (!includeForThisMailbox)
                                continue;

                            int attendeeCount = 0;
                            if (ev["attendees"] is JArray arr) attendeeCount = arr.Count;

                            var subjectStr = ev.SelectToken("subject")?.ToString();
                            if (string.IsNullOrWhiteSpace(subjectStr)) subjectStr = "[No Title]";

                            var organizerName = ev.SelectToken("organizer.emailAddress.name")?.ToString();
                            var organizerEmail = ev.SelectToken("organizer.emailAddress.address")?.ToString();

                            var eventId = ev.SelectToken("id")?.ToString();
                            var icalUid = ev.SelectToken("iCalUId")?.ToString();

                            // Use the mailbox’s friendly room name as Location so each room row shows correctly
                            var effectiveLocation = thisMailboxRoomName
                                                    ?? roomHits.FirstOrDefault()
                                                    ?? primaryLocation
                                                    ?? "Unknown";

                            allMeetings.Add(new MeetingViewModel
                            {
                                Id = eventId,
                                ICalUId = icalUid,
                                Subject = subjectStr,
                                StartTime = startIst.ToString("yyyy-MM-dd'T'HH:mm:ss"),
                                EndTime = endIst.ToString("yyyy-MM-dd'T'HH:mm:ss"),
                                Organizer = organizerName,
                                OrganizerEmail = organizerEmail,
                                Location = effectiveLocation,
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

                // === MULTI-ROOM ENRICHMENT (group by iCalUId or fallback key) ===
                var list = allMeetings.ToList();

                string KeyFor(MeetingViewModel m) =>
                    (m.ICalUId ?? $"{m.Subject}|{m.StartTime}|{m.EndTime}|{m.OrganizerEmail}")
                        .ToLowerInvariant();

                var roomsByKey = list
                    .GroupBy(m => KeyFor(m))
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(x => x.Location)
                              .Where(s => !string.IsNullOrWhiteSpace(s))
                              .Distinct(StringComparer.OrdinalIgnoreCase)
                              .OrderBy(s => s)
                              .ToList()
                    );

                var enriched = list
                    .Select(m =>
                    {
                        var k = KeyFor(m);
                        var rooms = roomsByKey.TryGetValue(k, out var r)
                            ? r
                            : new List<string> { m.Location ?? "Unknown" };

                        return new
                        {
                            id = m.Id,
                            iCalUId = m.ICalUId,
                            subject = m.Subject,
                            startTime = m.StartTime,
                            endTime = m.EndTime,
                            organizer = m.Organizer,
                            organizerEmail = m.OrganizerEmail,
                            location = m.Location,
                            attendeeCount = m.AttendeeCount,
                            multiRooms = rooms,
                            multiRoomCount = rooms.Count,
                            multiRoom = rooms.Count > 1
                        };
                    })
                    .OrderBy(x => x.startTime, StringComparer.Ordinal)
                    .ToList();

                return Ok(new { status = "success", count = enriched.Count, meetings = enriched });
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

        [HttpDelete("by-ical/{icalUid}")]
        public async Task<IActionResult> DeleteMeeting(string icalUid, [FromQuery] string organizerEmail)
        {
            if (string.IsNullOrWhiteSpace(icalUid) || string.IsNullOrWhiteSpace(organizerEmail))
                return BadRequest(new { status = "failure", message = "icalUid and organizerEmail are required." });

            var jwt = Request.Headers["Authorization"].ToString().Replace("Bearer ", "");
            if (string.IsNullOrWhiteSpace(jwt))
                return Unauthorized(new { status = "failure", message = "Missing user access token." });

            using var httpClient = _httpClientFactory.CreateClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwt);

            var lookupUrl = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(organizerEmail)}/events?$filter=iCalUId eq '{icalUid}'";
            var lookupResp = await httpClient.GetAsync(lookupUrl);

            if (!lookupResp.IsSuccessStatusCode)
            {
                var lookupError = await lookupResp.Content.ReadAsStringAsync();
                return StatusCode((int)lookupResp.StatusCode, new { status = "failure", message = "Lookup failed", details = lookupError });
            }

            var lookupJson = JObject.Parse(await lookupResp.Content.ReadAsStringAsync());
            var organizerEventId = lookupJson["value"]?.FirstOrDefault()?["id"]?.ToString();

            if (string.IsNullOrEmpty(organizerEventId))
                return NotFound(new { status = "failure", message = "Event not found in organizer’s mailbox" });

            var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(organizerEmail)}/events/{Uri.EscapeDataString(organizerEventId)}";
            var deleteResp = await httpClient.DeleteAsync(deleteUrl);

            if (!deleteResp.IsSuccessStatusCode)
            {
                var errorBody = await deleteResp.Content.ReadAsStringAsync();
                return StatusCode((int)deleteResp.StatusCode, new { status = "failure", message = "Graph deletion failed", details = errorBody });
            }
            return Ok(new { status = "success", message = "Meeting cancelled successfully." });
        }
    }
}
