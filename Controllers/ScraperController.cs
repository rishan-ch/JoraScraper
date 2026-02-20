using JoraScraper.Modules.Scraper.Interface;
using Microsoft.AspNetCore.Mvc;

namespace JoraScraper.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ScraperController : ControllerBase
    {
        private readonly IScraperService _scraperService;
        private readonly ILogger<ScraperController> _logger;

        public ScraperController(IScraperService scraperService, ILogger<ScraperController> logger)
        {
            _scraperService = scraperService;
            _logger = logger;
        }

        /// <summary>
        /// Triggers the scraper in the background and immediately returns 202 Accepted.
        /// </summary>
        [HttpPost("trigger")]
        public IActionResult TriggerScrape()
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    _logger.LogInformation("Manual scrape triggered via API at: {time}", DateTime.Now);
                    await _scraperService.ScrapeAndSaveJobsAsync();
                    _logger.LogInformation("Manual scrape completed at: {time}", DateTime.Now);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error during manual scrape triggered via API.");
                }
            });

            return Accepted(new
            {
                success = true,
                message = "Scraping job has been accepted and is running in the background. Check back shortly and use /api/scraper/export to download the result."
            });
        }

        /// <summary>
        /// Downloads the latest scraped Excel file.
        /// Returns 404 if not yet generated, or 202 with a hint if the file is stale/missing.
        /// </summary>
        [HttpGet("export")]
        public IActionResult ExportJobsExcel()
        {
            try
            {
                var fileName = "JobsFromJora.xlsx";
                var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
                var filePath = Path.Combine(exportDir, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "Excel file not found. Please trigger the scraper first via POST /api/scraper/trigger."
                    });
                }

                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(
                    fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message = "Failed to export Excel file.",
                    error = ex.Message
                });
            }
        }

        /// <summary>
        /// Returns the status of the latest Excel export (exists, last modified, age).
        /// </summary>
        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            var fileName = "JobsFromJora.xlsx";
            var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
            var filePath = Path.Combine(exportDir, fileName);

            if (!System.IO.File.Exists(filePath))
            {
                return Ok(new
                {
                    fileExists = false,
                    message = "No export file found. Trigger a scrape via POST /api/scraper/trigger."
                });
            }

            var lastModified = System.IO.File.GetLastWriteTimeUtc(filePath);
            var age = DateTime.UtcNow - lastModified;

            return Ok(new
            {
                fileExists = true,
                lastModifiedUtc = lastModified,
                ageHours = Math.Round(age.TotalHours, 1),
                isStale = age.TotalDays >= 2,
                message = age.TotalDays >= 2
                    ? "File is stale (>= 2 days old). The background service will re-scrape soon, or trigger manually."
                    : "File is fresh."
            });
        }
    }
}