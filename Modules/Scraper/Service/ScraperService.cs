using JoraScraper.Modules.Scraper.Dto;
using JoraScraper.Modules.Scraper.Interface;
using System.Text.RegularExpressions;
using OfficeOpenXml;
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
                _logger.LogInformation("Starting ultra-light browser setup for Render...");

                var executablePath = Environment.GetEnvironmentVariable("PUPPETEER_EXECUTABLE_PATH");
                
                var launchOptions = new LaunchOptions
                {
                    Headless = true,
                    ExecutablePath = executablePath,
                    Args = new[] { 
                        "--no-sandbox", 
                        "--disable-setuid-sandbox", 
                        "--disable-dev-shm-usage",
                        "--disable-gpu",
                        "--no-first-run",
                        "--no-zygote",
                        "--single-process", // CRITICAL: Run in one process to save RAM
                        "--disable-extensions",
                        "--disable-notifications"
                    },
                    Timeout = 120000 // 2 minute timeout for slow startup
                };

                // Local fallback
                if (string.IsNullOrEmpty(executablePath))
                {
                    _logger.LogInformation("PUPPETEER_EXECUTABLE_PATH not set. Using local Chromium.");
                    await new BrowserFetcher().DownloadAsync();
                }

                await using var browser = await Puppeteer.LaunchAsync(launchOptions);
                await Task.Delay(2000); // Let browser process settle

                await using var page = await browser.NewPageAsync();

                // --- RAM SAVER: Block heavy assets ---
                await page.SetRequestInterceptionAsync(true);
                page.Request += async (sender, e) =>
                {
                    var type = e.Request.ResourceType;
                    if (type == ResourceType.Image || type == ResourceType.Font || type == ResourceType.StyleSheet)
                    {
                        await e.Request.AbortAsync();
                    }
                    else if (e.Request.Url.Contains("analytics") || e.Request.Url.Contains("facebook"))
                    {
                        await e.Request.AbortAsync();
                    }
                    else
                    {
                        await e.Request.ContinueAsync();
                    }
                };

                await page.SetViewportAsync(new ViewPortOptions { Width = 1024, Height = 768 });
                await page.SetUserAgentAsync("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36");

                string targetUrl = "https://au.jora.com/j?sp=homepage&q=&l=";
                _logger.LogInformation("Navigating to Jora...");

                await page.GoToAsync(targetUrl, new NavigationOptions { 
                    WaitUntil = new[] { WaitUntilNavigation.Networkidle2 }, 
                    Timeout = 120000 
                });

                // Wait for the main results container
                try 
                {
                    await page.WaitForSelectorAsync(".job-card", new WaitForSelectorOptions { Timeout = 30000 });
                }
                catch 
                {
                    _logger.LogWarning("Timed out waiting for .job-card. Site might be blocking or empty.");
                    return;
                }

                // Get job card count
                var jobCount = await page.EvaluateFunctionAsync<int>("() => document.querySelectorAll('.job-card').length");
                _logger.LogInformation("Found {count} job cards. Starting extraction...", jobCount);

                // Limit to 20 jobs to avoid OOM (Out of Memory) on Free Tier
                for (int i = 0; i < Math.Min(jobCount, 20); i++)
                {
                    try
                    {
                        var job = await page.EvaluateFunctionAsync<JobInfo>($@"(index) => {{
                            const cards = document.querySelectorAll('.job-card');
                            const card = cards[index];
                            if (!card) return null;

                            const link = card.querySelector('a.job-link');
                            const company = card.querySelector('.job-company');
                            const location = card.querySelector('.job-location');
                            const abstract = card.querySelector('.job-abstract');

                            return {{
                                title: link ? link.innerText.trim() : 'Unknown Title',
                                url: link ? link.href : '',
                                company: company ? company.innerText.trim() : 'Unknown Company',
                                location: location ? location.innerText.trim() : 'N/A',
                                shortDescription: abstract ? abstract.innerText.trim() : '',
                                postedDate: card.querySelector('.job-listed-date')?.innerText.trim() || ''
                            }};
                        }}", i);

                        if (job != null && !string.IsNullOrEmpty(job.Url))
                        {
                            job.Description = job.ShortDescription; // Skip clicking on Render to save RAM
                            jobs.Add(job);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Error extracting job {index}: {msg}", i, ex.Message);
                    }
                }

                _logger.LogInformation("Scraping complete. Saving {count} jobs to Excel...", jobs.Count);
                await SaveJobsToExcel(jobs);

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Critical failure in ScraperService.");
            }
        }

        public async Task SaveJobsToExcel(List<JobInfo> jobs, string fileName = "JobsFromJora.xlsx")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
            if (!Directory.Exists(exportDir)) Directory.CreateDirectory(exportDir);

            var filePath = Path.Combine(exportDir, fileName);
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Jobs");

            // Headers
            ws.Cells[1, 1].Value = "JobPostId";
            ws.Cells[1, 2].Value = "Title";
            ws.Cells[1, 3].Value = "Company";
            ws.Cells[1, 4].Value = "Location";
            ws.Cells[1, 5].Value = "PostedDate";
            ws.Cells[1, 6].Value = "URL";
            ws.Cells[1, 7].Value = "Description";

            for (int i = 0; i < jobs.Count; i++)
            {
                var job = jobs[i];
                ws.Cells[i + 2, 1].Value = job.JobPostId.ToString();
                ws.Cells[i + 2, 2].Value = job.Title;
                ws.Cells[i + 2, 3].Value = job.Company;
                ws.Cells[i + 2, 4].Value = job.Location;
                ws.Cells[i + 2, 5].Value = job.PostedDate;
                ws.Cells[i + 2, 6].Value = job.Url;
                ws.Cells[i + 2, 7].Value = job.Description;
            }

            ws.Cells.AutoFitColumns();
            await package.SaveAsAsync(new FileInfo(filePath));
            _logger.LogInformation("Excel file saved successfully at {path}", filePath);
        }
    }
}