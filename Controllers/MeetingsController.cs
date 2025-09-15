using System.Globalization;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System.Collections.Concurrent;
using Teams_Backend_API.Models.Entities;
using System.Linq;
using System.Collections.Generic;
using Microsoft.AspNetCore.Http;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using Teams_Backend_API.Models.DTOs;

namespace OutLook_Events
{
    [Route("api/[controller]")]
    [ApiController]
    public class MeetingsController : ControllerBase
    {
        private static readonly string OrgDomain = "conservesolution.com";

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

        private IConfidentialClientApplication BuildApp()
        {
            var clientId = _config["AzureAd:ClientId"];
            var clientSecret = _config["AzureAd:ClientSecret"];
            var tenantId = _config["AzureAd:TenantId"];

            if (string.IsNullOrWhiteSpace(clientId) ||
                string.IsNullOrWhiteSpace(clientSecret) ||
                string.IsNullOrWhiteSpace(tenantId))
            {
                throw new ApplicationException("Azure AD configuration missing (ClientId/ClientSecret/TenantId).");
            }

            return ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .Build();
        }

        private static string? GetEmailFromJwt(string jwt)
        {
            try
            {
                var handler = new JwtSecurityTokenHandler();
                var token = handler.ReadJwtToken(jwt);
                var email = token.Claims.FirstOrDefault(c =>
                                c.Type.Equals("preferred_username", StringComparison.OrdinalIgnoreCase) ||
                                c.Type.Equals("upn", StringComparison.OrdinalIgnoreCase) ||
                                c.Type.Equals(ClaimTypes.Upn, StringComparison.OrdinalIgnoreCase) ||
                                c.Type.Equals(ClaimTypes.Email, StringComparison.OrdinalIgnoreCase) ||
                                c.Type.Equals("unique_name", StringComparison.OrdinalIgnoreCase))
                            ?.Value;
                return email?.ToLowerInvariant();
            }
            catch { return null; }
        }

        // Dynamic admin check using Graph
        private async Task<bool> IsUserAdminAsync(string jwt)
        {
            try
            {
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwt);

                // Get signed-in user
                var meResp = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/me");
                if (!meResp.IsSuccessStatusCode) return false;

                var meJson = JObject.Parse(await meResp.Content.ReadAsStringAsync());
                var userId = meJson["id"]?.ToString();
                if (string.IsNullOrEmpty(userId)) return false;

                // Get role memberships
                var roleResp = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/users/{userId}/memberOf");
                if (!roleResp.IsSuccessStatusCode) return false;

                var roleJson = JObject.Parse(await roleResp.Content.ReadAsStringAsync());
                var roles = roleJson["value"]?.Select(r => r["displayName"]?.ToString() ?? "").ToList();

                return roles != null && roles.Any(r =>
                    r.Equals("Company Administrator", StringComparison.OrdinalIgnoreCase) ||
                    r.Equals("Global Administrator", StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to check admin role.");
                return false;
            }
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

            var mailboxToRoomName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["gfmeeting@conservesolution.com"] = "Ground Floor Meeting Room",
                ["ffmeeting@conservesolution.com"] = "1st Floor Meeting Room",
                ["contconference@conservesolution.com"] = "Conference Room",
                ["sfmeeting@conservesolution.com"] = "3rd Floor Meeting Room",
            };

            var allowedLocations = new HashSet<string>(mailboxToRoomName.Values, StringComparer.OrdinalIgnoreCase);

            try
            {
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                var app = BuildApp();
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

                        mailboxToRoomName.TryGetValue(email, out var thisMailboxRoomName);

                        foreach (var ev in events)
                        {
                            var startStr = ev.SelectToken("start.dateTime")?.ToString();
                            var endStr = ev.SelectToken("end.dateTime")?.ToString();

                            if (!DateTime.TryParse(startStr, out var startIst) ||
                                !DateTime.TryParse(endStr, out var endIst))
                                continue;

                            if (startIst.Date != selectedDateIst) continue;

                            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            var primaryLocation = ev.SelectToken("location.displayName")?.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(primaryLocation)) names.Add(primaryLocation);

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

                            var roomHits = names.Where(n => allowedLocations.Contains(n)).ToList();
                            bool includeForThisMailbox =
                                (!string.IsNullOrEmpty(thisMailboxRoomName) && roomHits.Contains(thisMailboxRoomName))
                                || (string.IsNullOrEmpty(thisMailboxRoomName) && roomHits.Count > 0)
                                || (roomHits.Count == 0 && !string.IsNullOrEmpty(thisMailboxRoomName));

                            if (!includeForThisMailbox) continue;

                            int attendeeCount = 0;
                            if (ev["attendees"] is JArray arr) attendeeCount = arr.Count;

                            var subjectStr = ev.SelectToken("subject")?.ToString();
                            if (string.IsNullOrWhiteSpace(subjectStr)) subjectStr = "[No Title]";

                            var organizerName = ev.SelectToken("organizer.emailAddress.name")?.ToString();
                            var organizerEmail = ev.SelectToken("organizer.emailAddress.address")?.ToString();

                            var eventId = ev.SelectToken("id")?.ToString();
                            var icalUid = ev.SelectToken("iCalUId")?.ToString();

                            var effectiveLocation =
                                thisMailboxRoomName
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

                var list = allMeetings.ToList();

                string KeyFor(MeetingViewModel m) =>
                    (m.ICalUId ?? $"{m.Subject}|{m.StartTime}|{m.EndTime}|{m.OrganizerEmail}").ToLowerInvariant();

                var roomsByKey = list
                    .GroupBy(m => KeyFor(m))
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(x => x.Location)
                              .Where(s => !string.IsNullOrWhiteSpace(s))
                              .Select(s => s!)
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

        private async Task<(bool ok, string? eventId, string? error, int status)> FindOrganizerEventId(HttpClient httpClient, string organizerEmail, string icalUid)
        {
            var lookupUrl = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(organizerEmail)}/events?$filter=iCalUId eq '{icalUid}'";
            var lookupResp = await httpClient.GetAsync(lookupUrl);
            if (!lookupResp.IsSuccessStatusCode)
            {
                var body = await lookupResp.Content.ReadAsStringAsync();
                return (false, null, body, (int)lookupResp.StatusCode);
            }
            var lookupJson = JObject.Parse(await lookupResp.Content.ReadAsStringAsync());
            var organizerEventId = lookupJson["value"]?.FirstOrDefault()?["id"]?.ToString();
            if (string.IsNullOrEmpty(organizerEventId))
                return (false, null, "Event not found in organizer's mailbox", 404);
            return (true, organizerEventId, null, 200);
        }

        private static bool IsCompleted(string startIso, string endIso)
        {
            if (!DateTime.TryParse(startIso, out var s)) return false;
            if (!DateTime.TryParse(endIso, out var e)) return false;
            return DateTime.UtcNow > e.ToUniversalTime();
        }

        [HttpDelete("by-ical/{icalUid}")]
        public async Task<IActionResult> DeleteMeeting(string icalUid, [FromQuery] string organizerEmail)
        {
            if (string.IsNullOrWhiteSpace(icalUid) || string.IsNullOrWhiteSpace(organizerEmail))
                return BadRequest(new { status = "failure", message = "icalUid and organizerEmail are required." });

            var jwt = Request.Headers["Authorization"].ToString().Replace("Bearer ", "");
            if (string.IsNullOrWhiteSpace(jwt))
                return Unauthorized(new { status = "failure", message = "Missing user access token." });

            var callerEmail = GetEmailFromJwt(jwt);
            var isAdmin = await IsUserAdminAsync(jwt);
            var isOrgMail = callerEmail != null && callerEmail.EndsWith($"@{OrgDomain}", StringComparison.OrdinalIgnoreCase);

            if (!isAdmin && !isOrgMail)
                return Unauthorized(new { status = "failure", message = $"Please sign in with your @{OrgDomain} account." });

            try
            {
                HttpClient httpClient = _httpClientFactory.CreateClient();
                if (isAdmin)
                {
                    var app = BuildApp();
                    var scopes = new[] { "https://graph.microsoft.com/.default" };
                    var appToken = await app.AcquireTokenForClient(scopes).ExecuteAsync();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", appToken.AccessToken);
                }
                else
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwt);
                }

                var (ok, eventId, error, status) = await FindOrganizerEventId(httpClient, organizerEmail, icalUid);
                if (!ok) return StatusCode(status, new { status = "failure", message = "Lookup failed", details = error });

                if (!isAdmin && !string.Equals(callerEmail, organizerEmail, StringComparison.OrdinalIgnoreCase))
                    return Unauthorized(new { status = "failure", message = "Only the organizer or an admin can delete this meeting." });

                var evUrl = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(organizerEmail)}/events/{Uri.EscapeDataString(eventId!)}?$select=start,end";
                var evResp = await httpClient.GetAsync(evUrl);
                if (evResp.IsSuccessStatusCode)
                {
                    var obj = JObject.Parse(await evResp.Content.ReadAsStringAsync());
                    var startIso = obj.SelectToken("start.dateTime")?.ToString() ?? "";
                    var endIso = obj.SelectToken("end.dateTime")?.ToString() ?? "";
                    if (!isAdmin && IsCompleted(startIso, endIso))
                        return Unauthorized(new { status = "failure", message = "Completed meetings cannot be deleted by non-admins." });
                }

                var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(organizerEmail)}/events/{Uri.EscapeDataString(eventId!)}";
                var deleteResp = await httpClient.DeleteAsync(deleteUrl);

                if (!deleteResp.IsSuccessStatusCode)
                {
                    var errorBody = await deleteResp.Content.ReadAsStringAsync();
                    return StatusCode((int)deleteResp.StatusCode, new { status = "failure", message = "Graph deletion failed", details = errorBody });
                }

                return Ok(new { status = "success", message = "Meeting cancelled successfully." });
            }
            catch (MsalServiceException msalEx)
            {
                _logger.LogError(msalEx, "Azure AD token acquisition failed (admin override).");
                return StatusCode(500, new { status = "failure", message = "Azure AD token acquisition failed." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in DeleteMeeting");
                return StatusCode(500, new { status = "failure", message = "Internal server error", details = ex.Message });
            }
        }

        [HttpPatch("by-ical/{icalUid}")]
        public async Task<IActionResult> PatchMeeting(string icalUid, [FromBody] PatchMeetingDto dto)
        {
            if (string.IsNullOrWhiteSpace(icalUid))
                return BadRequest(new { status = "failure", message = "icalUid is required." });
            if (dto == null || string.IsNullOrWhiteSpace(dto.OrganizerEmail))
                return BadRequest(new { status = "failure", message = "OrganizerEmail is required." });
            if (dto.Subject != null && dto.Subject.Trim().Length < 3)
                return BadRequest(new { status = "failure", message = "Subject must be at least 3 characters." });

            var jwt = Request.Headers["Authorization"].ToString().Replace("Bearer ", "");
            if (string.IsNullOrWhiteSpace(jwt))
                return Unauthorized(new { status = "failure", message = "Missing user access token." });

            var callerEmail = GetEmailFromJwt(jwt);
            var isAdmin = await IsUserAdminAsync(jwt);
            var isOrgMail = callerEmail != null && callerEmail.EndsWith($"@{OrgDomain}", StringComparison.OrdinalIgnoreCase);

            if (!isAdmin && !isOrgMail)
                return Unauthorized(new { status = "failure", message = $"Please sign in with your @{OrgDomain} account." });

            try
            {
                HttpClient httpClient = _httpClientFactory.CreateClient();
                if (isAdmin)
                {
                    var app = BuildApp();
                    var scopes = new[] { "https://graph.microsoft.com/.default" };
                    var appToken = await app.AcquireTokenForClient(scopes).ExecuteAsync();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", appToken.AccessToken);
                }
                else
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwt);
                }

                var (ok, eventId, error, status) = await FindOrganizerEventId(httpClient, dto.OrganizerEmail, icalUid);
                if (!ok) return StatusCode(status, new { status = "failure", message = "Lookup failed", details = error });

                if (!isAdmin && !string.Equals(callerEmail, dto.OrganizerEmail, StringComparison.OrdinalIgnoreCase))
                    return Unauthorized(new { status = "failure", message = "Only the organizer or an admin can edit this meeting." });

                var patchBody = new JObject();
                if (!string.IsNullOrWhiteSpace(dto.Subject)) patchBody["subject"] = dto.Subject.Trim();
                if (!string.IsNullOrWhiteSpace(dto.StartTime) && !string.IsNullOrWhiteSpace(dto.EndTime))
                {
                    patchBody["start"] = new JObject { ["dateTime"] = dto.StartTime, ["timeZone"] = "India Standard Time" };
                    patchBody["end"] = new JObject { ["dateTime"] = dto.EndTime, ["timeZone"] = "India Standard Time" };
                }

                if (!patchBody.HasValues)
                    return BadRequest(new { status = "failure", message = "No changes to apply." });

                var req = new HttpRequestMessage(new HttpMethod("PATCH"),
                    $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(dto.OrganizerEmail)}/events/{Uri.EscapeDataString(eventId!)}")
                {
                    Content = new StringContent(patchBody.ToString(), System.Text.Encoding.UTF8, "application/json")
                };

                var resp = await httpClient.SendAsync(req);
                if (!resp.IsSuccessStatusCode)
                {
                    var body = await resp.Content.ReadAsStringAsync();
                    return StatusCode((int)resp.StatusCode, new { status = "failure", message = "Graph update failed", details = body });
                }

                return Ok(new { status = "success", message = "Meeting updated." });
            }
            catch (MsalServiceException msalEx)
            {
                _logger.LogError(msalEx, "Azure AD token acquisition failed (admin override).");
                return StatusCode(500, new { status = "failure", message = "Azure AD token acquisition failed." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in PatchMeeting");
                return StatusCode(500, new { status = "failure", message = "Internal server error", details = ex.Message });
            }
        }
    }
}