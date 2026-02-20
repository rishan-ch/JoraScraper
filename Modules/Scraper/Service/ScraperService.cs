using JoraScraper.Modules.Scraper.Dto;
using JoraScraper.Modules.Scraper.Interface;
using System.Text.RegularExpressions;
using OfficeOpenXml; // Required for ExcelPackage
using PuppeteerSharp;

namespace JoraScraper.Modules.Scraper.Service
{
    public class ScraperService : IScraperService
    {
        private readonly ILogger<ScraperService> _logger;

        public ScraperService(ILogger<ScraperService> logger)
        {
            _logger = logger;
        }

        public async Task ScrapeAndSaveJobsAsync()
        {
            var jobs = new List<JobInfo>();
            var processedUrls = new HashSet<string>();

            try
            {
                _logger.LogInformation("Starting browser setup...");
                var executablePath = Environment.GetEnvironmentVariable("PUPPETEER_EXECUTABLE_PATH");
                LaunchOptions launchOptions;

                if (!string.IsNullOrEmpty(executablePath))
                {
                    _logger.LogInformation("Using system Chromium at: {path}", executablePath);
                    launchOptions = new LaunchOptions
                    {
                        Headless = true,
                        ExecutablePath = executablePath,
                        Args = new[] { "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage" }
                    };
                }
                else
                {
                    _logger.LogInformation("PUPPETEER_EXECUTABLE_PATH not set. Downloading Chromium...");
                    var browserFetcher = new BrowserFetcher();
                    await browserFetcher.DownloadAsync();
                    launchOptions = new LaunchOptions
                    {
                        Headless = true,
                        Args = new[] { "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage" }
                    };
                }

                await using var browser = await Puppeteer.LaunchAsync(launchOptions);
                await using var page = await browser.NewPageAsync();

                await page.SetViewportAsync(new ViewPortOptions { Width = 1920, Height = 1080 });
                await page.SetUserAgentAsync("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36");
                
                // Navigation and Scraping Logic
                string firstUrl = "https://au.jora.com/j?sp=homepage&trigger_source=homepage&q=&l=";
                _logger.LogInformation("Navigating to Jora...");
                await page.GoToAsync(firstUrl, new NavigationOptions { WaitUntil = new[] { WaitUntilNavigation.DOMContentLoaded }, Timeout = 30000 });

                // ... (Your existing scraping loop that found 15 jobs goes here)
                // For brevity, I'm skipping to the Save call.
                
                _logger.LogInformation("Scraping complete. Total jobs collected: {count}", jobs.Count);
                await SaveJobsToExcel(jobs);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during scraping process.");
            }
        }

        public async Task SaveJobsToExcel(List<JobInfo> jobs, string fileName = "JobsFromJora.xlsx")
        {
            // --- THE CRITICAL FIX START ---
            // Set the license using the NEW static property. 
            // DO NOT use 'ExcelPackage.LicenseContext'.
            ExcelPackage.License.SetNonCommercialPersonal("Your Name");
            // --- THE CRITICAL FIX END ---

            var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
            if (!Directory.Exists(exportDir))
                Directory.CreateDirectory(exportDir);

            var filePath = Path.Combine(exportDir, fileName);

            if (File.Exists(filePath))
                File.Delete(filePath);

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Jobs");

            // Headers
            string[] headers = { "JobPostId", "Title", "Company", "Location", "Salary", "Posted Date", "URL", "Description" };
            for (int h = 0; h < headers.Length; h++)
            {
                worksheet.Cells[1, h + 1].Value = headers[h];
                worksheet.Cells[1, h + 1].Style.Font.Bold = true;
            }

            // Data rows
            for (int i = 0; i < jobs.Count; i++)
            {
                var job = jobs[i];
                worksheet.Cells[i + 2, 1].Value = job.JobPostId;
                worksheet.Cells[i + 2, 2].Value = job.Title;
                worksheet.Cells[i + 2, 3].Value = job.Company;
                worksheet.Cells[i + 2, 4].Value = job.Location;
                worksheet.Cells[i + 2, 5].Value = job.Salary;
                worksheet.Cells[i + 2, 6].Value = job.PostedDate;
                worksheet.Cells[i + 2, 7].Value = job.Url;
                worksheet.Cells[i + 2, 8].Value = job.DescriptionHtml;
            }

            worksheet.Cells.AutoFitColumns();
            await package.SaveAsAsync(new FileInfo(filePath));

            _logger.LogInformation("✅ Excel saved successfully to: {filePath}", filePath);
        }

        private string CleanText(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = Regex.Replace(text, @"[ \t]+", " ");
            text = Regex.Replace(text, @"\n\s*\n\s*\n+", "\n\n");
            return text.Trim();
        }

        // Inner class for description extraction
        private class DescriptionData { public string Text { get; set; } = ""; public string Html { get; set; } = ""; }
    }
}