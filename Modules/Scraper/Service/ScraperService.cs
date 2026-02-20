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
                _logger.LogInformation("Starting browser setup...");

                var executablePath = Environment.GetEnvironmentVariable("PUPPETEER_EXECUTABLE_PATH");
                
                // Optimized Launch Options for low-resource environments (Render Free Tier)
                var launchOptions = new LaunchOptions
                {
                    Headless = true,
                    ExecutablePath = executablePath,
                    Args = new[] { 
                        "--no-sandbox", 
                        "--disable-setuid-sandbox", 
                        "--disable-dev-shm-usage",
                        "--disable-gpu",            // Saves RAM
                        "--disable-extensions",     // Saves RAM
                        "--no-zygote",              // Reduces process overhead
                        "--single-process",         // Essential for 512MB RAM limits
                        "--no-first-run"
                    },
                    Timeout = 60000 // Increase launch timeout to 60s
                };

                if (string.IsNullOrEmpty(executablePath))
                {
                    _logger.LogInformation("Downloading Chromium for local development...");
                    await new BrowserFetcher().DownloadAsync();
                }

                _logger.LogInformation("Launching Browser instance...");
                await using var browser = await Puppeteer.LaunchAsync(launchOptions);
                
                // Give the browser process a moment to stabilize on slow CPUs
                await Task.Delay(2000);

                await using var page = await browser.NewPageAsync();
                
                // Set a realistic identity
                await page.SetViewportAsync(new ViewPortOptions { Width = 1280, Height = 800 });
                await page.SetUserAgentAsync("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36");

                string firstUrl = "https://au.jora.com/j?sp=homepage&q=&l=";
                _logger.LogInformation("Navigating to Jora...");

                // Use Networkidle2 for better stability on slow loads
                await page.GoToAsync(firstUrl, new NavigationOptions { 
                    WaitUntil = new[] { WaitUntilNavigation.Networkidle2 }, 
                    Timeout = 90000 
                });

                int pageNum = 1;
                int totalPages = 1;

                try
                {
                    await page.WaitForSelectorAsync(".job-card.result", new WaitForSelectorOptions { Timeout = 15000 });
                    totalPages = await GetTotalPages(page);
                    _logger.LogInformation("Detected {totalPages} pages.", totalPages);
                }
                catch
                {
                    _logger.LogWarning("Job cards not found or page took too long to render.");
                }

                // Limit pages for free tier to prevent OOM (Out of Memory)
                int maxPages = 3; 

                while (pageNum <= totalPages && pageNum <= maxPages)
                {
                    string url = pageNum == 1 ? firstUrl : $"{firstUrl}&p={pageNum}";
                    
                    if (pageNum > 1)
                    {
                        _logger.LogInformation("Moving to page {pageNum}...", pageNum);
                        await page.GoToAsync(url, new NavigationOptions { WaitUntil = new[] { WaitUntilNavigation.Networkidle2 } });
                    }

                    await Task.Delay(3000); // Anti-bot breath

                    var jobCardsCount = await page.EvaluateFunctionAsync<int>("() => document.querySelectorAll('.job-card.result').length");
                    
                    for (int i = 0; i < jobCardsCount; i++)
                    {
                        try 
                        {
                            var jobInfo = await ExtractJobBasicInfo(page, i);
                            
                            if (jobInfo != null && !processedUrls.Contains(jobInfo.Url))
                            {
                                processedUrls.Add(jobInfo.Url);
                                
                                // Fetch Description - Note: Click-to-expand is risky on low RAM
                                // We'll try it, but wrap it tightly in a try-catch
                                try 
                                {
                                    var clickSelector = $".job-card.result:nth-of-type({i + 1}) a.job-link.show-job-description";
                                    await page.ClickAsync(clickSelector);
                                    await page.WaitForSelectorAsync(".job-description-container", new WaitForSelectorOptions { Timeout = 5000 });
                                    
                                    var descData = await page.EvaluateFunctionAsync<DescriptionData>(@"() => {
                                        const container = document.querySelector('.job-description-container');
                                        return {
                                            text: container ? container.innerText.trim() : '',
                                            html: container ? container.innerHTML.trim() : ''
                                        };
                                    }");
                                    jobInfo.Description = CleanText(descData.text);
                                }
                                catch { jobInfo.Description = "Description could not be loaded."; }

                                jobs.Add(jobInfo);
                            }
                        }
                        catch (Exception ex) { _logger.LogWarning("Skipping a job due to error: {msg}", ex.Message); }
                    }
                    pageNum++;
                }

                await SaveJobsToExcel(jobs);
                _logger.LogInformation("Scraping completed successfully.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Critical Error during scraping process.");
            }
        }

        private async Task<JobInfo> ExtractJobBasicInfo(IPage page, int index)
        {
            return await page.EvaluateFunctionAsync<JobInfo>($@"(index) => {{
                const card = document.querySelectorAll('.job-card.result')[index];
                if (!card) return null;
                const link = card.querySelector('a.job-link');
                return {{
                    title: link ? link.innerText.trim() : 'Unknown',
                    url: link ? link.href : '',
                    company: card.querySelector('.job-company')?.innerText.trim() || '',
                    location: card.querySelector('.job-location')?.innerText.trim() || '',
                    salary: card.querySelector('.badge')?.innerText.trim() || 'Not Specified',
                    postedDate: card.querySelector('.job-listed-date')?.innerText.trim() || ''
                }};
            }}", index);
        }

        private async Task<int> GetTotalPages(IPage page)
        {
            return await page.EvaluateFunctionAsync<int>(@"() => {
                const el = document.querySelector('.search-results-page-number');
                const match = el?.innerText.match(/of\s+(\d+)/i);
                return match ? parseInt(match[1]) : 1;
            }");
        }

        private string CleanText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return Regex.Replace(text, @"\s+", " ").Trim();
        }

        public async Task SaveJobsToExcel(List<JobInfo> jobs, string fileName = "JobsFromJora.xlsx")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
            if (!Directory.Exists(exportDir)) Directory.CreateDirectory(exportDir);

            var filePath = Path.Combine(exportDir, fileName);
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Jobs");
            
            // Header
            string[] headers = { "JobPostId", "Title", "Company", "Location", "Salary", "PostedDate", "URL", "Description" };
            for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

            for (int i = 0; i < jobs.Count; i++)
            {
                ws.Cells[i + 2, 1].Value = jobs[i].JobPostId.ToString();
                ws.Cells[i + 2, 2].Value = jobs[i].Title;
                ws.Cells[i + 2, 3].Value = jobs[i].Company;
                ws.Cells[i + 2, 4].Value = jobs[i].Location;
                ws.Cells[i + 2, 5].Value = jobs[i].Salary;
                ws.Cells[i + 2, 6].Value = jobs[i].PostedDate;
                ws.Cells[i + 2, 7].Value = jobs[i].Url;
                ws.Cells[i + 2, 8].Value = jobs[i].Description;
            }

            await package.SaveAsAsync(new FileInfo(filePath));
        }
    }

    public class DescriptionData { public string text { get; set; } public string html { get; set; } }
}