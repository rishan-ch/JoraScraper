using JoraScraper.Modules.Scraper.Interface;

namespace JoraScraper.Modules.Scraper.Service
{
    public class JobScraperBackgroundService : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<JobScraperBackgroundService> _logger;

        private static readonly string ExportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
        private static readonly string FilePath = Path.Combine(ExportDir, "JobsFromJora.xlsx");

        public JobScraperBackgroundService(IServiceProvider serviceProvider, ILogger<JobScraperBackgroundService> logger)
        {
            _serviceProvider = serviceProvider;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("JobScraperBackgroundService started at: {time}", DateTime.Now);

            while (!stoppingToken.IsCancellationRequested)
            {
                bool fileExists = File.Exists(FilePath);
                bool fileIsStale = fileExists && (DateTime.UtcNow - File.GetLastWriteTimeUtc(FilePath)).TotalDays >= 2;

                if (!fileExists || fileIsStale)
                {
                    _logger.LogInformation(
                        fileExists
                            ? "Excel file is stale (>= 2 days old). Starting scrape at: {time}"
                            : "Excel file not found. Starting scrape at: {time}",
                        DateTime.Now);

                    try
                    {
                        using var scope = _serviceProvider.CreateScope();
                        var scraperService = scope.ServiceProvider.GetRequiredService<IScraperService>();
                        await scraperService.ScrapeAndSaveJobsAsync();
                        _logger.LogInformation("Job scraping completed successfully at: {time}", DateTime.Now);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error occurred while scraping jobs at: {time}", DateTime.Now);
                    }
                }
                else
                {
                    _logger.LogInformation("Excel file is fresh (< 2 days old). Skipping scrape.");
                }

                // Check again every hour; actual scrape only happens when file is missing/stale
                _logger.LogInformation("Background service sleeping for 1 hour before next check...");
                await Task.Delay(TimeSpan.FromHours(1), stoppingToken);
            }

            _logger.LogInformation("JobScraperBackgroundService stopped at: {time}", DateTime.Now);
        }
    }
}