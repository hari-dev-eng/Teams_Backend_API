using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// Controllers & Swagger
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// CORS (local React + deployed Vercel origins)
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowReactApp", policy =>
    {
        policy.SetIsOriginAllowed(origin =>
        {
            if (string.IsNullOrEmpty(origin)) return false;
            var host = new Uri(origin).Host;
            return host.Equals("teams-meeting-web.vercel.app", StringComparison.OrdinalIgnoreCase)
                   || host.EndsWith(".vercel.app", StringComparison.OrdinalIgnoreCase)  // previews
                   || host.StartsWith("localhost");
        })
        .AllowAnyHeader()
        .AllowAnyMethod()
        .SetPreflightMaxAge(TimeSpan.FromHours(12));
    });
});

builder.Services.AddHttpClient();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// Forwarded Headers (MUST BE FIRST!)
var fwd = new ForwardedHeadersOptions
{
    ForwardedHeaders = ForwardedHeaders.XForwardedProto | ForwardedHeaders.XForwardedFor
};
fwd.KnownProxies.Clear();
fwd.KnownNetworks.Clear();
app.UseForwardedHeaders(fwd);  // Only ONCE, first

app.UseCors("AllowReactApp");
app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();
app.MapGet("/healthz", () => Results.Ok(new { status = "ok", timeUtc = DateTime.UtcNow }));
app.Run();
