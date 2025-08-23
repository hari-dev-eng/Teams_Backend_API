var builder = WebApplication.CreateBuilder(args);

// Add services
builder.Services.AddControllers();

// Add Swagger
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Add CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowReactApp", policy =>
    {
        builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowReactApp", policy =>
    {
        policy.WithOrigins(
                "http://localhost:3000",
                "http://localhost:3001",
                "http://localhost:3002",
                "https://teams-meeting-6nqqa3u12-hari-dev-engs-projects.vercel.app"
            )
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

builder.Services.AddHttpClient();

var app = builder.Build();

// Enable Swagger middleware
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

// Use the named CORS policy
app.UseCors("AllowReactApp");

app.UseAuthorization();
app.MapControllers();

app.Run();
