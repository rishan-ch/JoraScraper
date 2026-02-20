using DotNetEnv;
using JoraScraper.Modules.Scraper.Interface;
using JoraScraper.Modules.Scraper.Service;
using System.Text.Json.Serialization;

Env.Load();

var builder = WebApplication.CreateBuilder(args);

OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Your Name");

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowFrontend", policy =>
    {
        policy.WithOrigins(
                "https://www.dsailorgroup.com.au",
                "https://dsailor-vercel.vercel.app",
                "https://dsailorgroup.comm.au",
                "https://localhost:5193"
            )
            .AllowAnyMethod()
            .AllowAnyHeader()
            .AllowCredentials();
    });
});

builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
    });

builder.Services.AddControllersWithViews();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Scraper service (scoped — creates its own browser per scrape run)
builder.Services.AddScoped<IScraperService, ScraperService>();

// Background service — runs on startup, checks every hour, scrapes every 2 days
builder.Services.AddHostedService<JobScraperBackgroundService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// Do NOT use HTTPS redirection on Render — Render handles TLS termination externally
// app.UseHttpsRedirection();

app.UseCors("AllowFrontend");
app.UseAuthorization();
app.MapControllers();

app.Run();