using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Identity.Web;

var builder = WebApplication.CreateBuilder(args);

// Add configuration
builder.Configuration.AddEnvironmentVariables();

// Auth
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApi(options =>
    {
        builder.Configuration.Bind("AzureAd", options);
        options.TokenValidationParameters.ValidateAudience = true;
        options.TokenValidationParameters.ValidAudiences = new[]
        {
            builder.Configuration["AzureAd:ClientId"],
            $"api://{builder.Configuration["AzureAd:ClientId"]}"
        };
    },
    options =>
    {
        builder.Configuration.Bind("AzureAd", options);
    });

builder.Services.AddAuthorization();

// Controllers & Swagger
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title = "Teams Booking API",
        Version = "v1",
        Description = "API for managing Teams meeting room bookings"
    });

    // Add security definition for Swagger
    c.AddSecurityDefinition("Bearer", new Microsoft.OpenApi.Models.OpenApiSecurityScheme
    {
        Name = "Authorization",
        Type = Microsoft.OpenApi.Models.SecuritySchemeType.Http,
        Scheme = "Bearer",
        BearerFormat = "JWT",
        In = Microsoft.OpenApi.Models.ParameterLocation.Header,
        Description = "JWT Authorization header using the Bearer scheme."
    });

    c.AddSecurityRequirement(new Microsoft.OpenApi.Models.OpenApiSecurityRequirement
    {
        {
            new Microsoft.OpenApi.Models.OpenApiSecurityScheme
            {
                Reference = new Microsoft.OpenApi.Models.OpenApiReference
                {
                    Type = Microsoft.OpenApi.Models.ReferenceType.SecurityScheme,
                    Id = "Bearer"
                }
            },
            Array.Empty<string>()
        }
    });
});

// CORS - Production only configuration
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowProduction", policy =>
    {
        policy.WithOrigins("https://teams-meeting-web.vercel.app")
              .AllowAnyHeader()
              .AllowAnyMethod()
              .AllowCredentials();
    });

    // Add a more restrictive policy for production
    options.AddPolicy("StrictProduction", policy =>
    {
        policy.WithOrigins("https://teams-meeting-web.vercel.app")
              .WithHeaders("Content-Type", "Authorization", "Accept")
              .WithMethods("GET", "POST", "PUT", "DELETE", "OPTIONS")
              .AllowCredentials();
    });
});

// HTTP Client and Services
builder.Services.AddHttpClient();
builder.Services.AddHealthChecks();

// Add response compression for production
builder.Services.AddResponseCompression(options =>
{
    options.EnableForHttps = true;
});

var app = builder.Build();

// Security headers
app.Use(async (context, next) =>
{
    context.Response.Headers.Add("X-Content-Type-Options", "nosniff");
    context.Response.Headers.Add("X-Frame-Options", "DENY");
    context.Response.Headers.Add("X-XSS-Protection", "1; mode=block");
    context.Response.Headers.Add("Referrer-Policy", "strict-origin-when-cross-origin");

    await next();
});

// Response compression
app.UseResponseCompression();

// Enable Forwarded headers for production (behind proxy)
if (app.Environment.IsProduction())
{
    app.UseForwardedHeaders(new ForwardedHeadersOptions
    {
        ForwardedHeaders = ForwardedHeaders.XForwardedProto | ForwardedHeaders.XForwardedFor,
        ForwardLimit = 2
    });

    app.UseHttpsRedirection();
}

// Enable Swagger in Development only
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "Teams Booking API v1");
        c.RoutePrefix = "swagger";
    });
}

// CRITICAL: Middleware order must be correct
app.UseRouting();

// Use CORS - Must come before Authentication
app.UseCors("AllowProduction"); // Use "StrictProduction" for more security

app.UseAuthentication();
app.UseAuthorization();

// Endpoint routing
app.UseEndpoints(endpoints =>
{
    endpoints.MapControllers().RequireCors("AllowProduction");
    endpoints.MapHealthChecks("/healthz");
});

// Global error handling
app.UseExceptionHandler("/error");

// Add a simple error endpoint
app.Map("/error", () => Results.Problem("An unexpected error occurred."));

// Health check endpoint
app.MapGet("/healthz", () => Results.Ok(new
{
    status = "healthy",
    timestamp = DateTime.UtcNow,
    environment = app.Environment.EnvironmentName
}));

// Add API info endpoint
app.MapGet("/api/info", () => new
{
    apiVersion = "1.0",
    environment = app.Environment.EnvironmentName,
    supportedEndpoints = new[] { "/api/Bookings", "/healthz", "/api/info" },
    timestamp = DateTime.UtcNow
});

app.Run();