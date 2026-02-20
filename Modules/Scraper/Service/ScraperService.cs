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

                // Use system Chromium if PUPPETEER_EXECUTABLE_PATH is set (e.g. in Docker on Render),
                // otherwise fall back to downloading Chromium locally (local dev).
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
                await page.SetExtraHttpHeadersAsync(new Dictionary<string, string>
                {
                    { "Accept-Language", "en-US,en;q=0.9" },
                    { "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8" },
                    { "Sec-Fetch-Site", "none" },
                    { "Sec-Fetch-Mode", "navigate" },
                    { "Sec-Fetch-User", "?1" },
                    { "Sec-Fetch-Dest", "document" }
                });

                int pageNum = 1;
                int totalPages = 1;

                string firstUrl = "https://au.jora.com/j?sp=homepage&trigger_source=homepage&q=&l=";
                _logger.LogInformation("Navigating to first page...");

                await page.GoToAsync(firstUrl, new NavigationOptions { WaitUntil = new[] { WaitUntilNavigation.DOMContentLoaded }, Timeout = 30000 });

                try
                {
                    await page.WaitForSelectorAsync(".job-card.result", new WaitForSelectorOptions { Timeout = 10000 });
                    totalPages = await GetTotalPages(page);
                    _logger.LogInformation("Detected total pages: {totalPages}", totalPages);
                }
                catch
                {
                    _logger.LogWarning("No job cards found on first page.");
                }

                while (pageNum <= totalPages)
                {
                    string url = pageNum == 1
                        ? firstUrl
                        : $"https://au.jora.com/j?sp=homepage&trigger_source=homepage&q=&l=&p={pageNum}";

                    _logger.LogInformation("Processing page {pageNum}/{totalPages}: {url}", pageNum, totalPages, url);

                    if (pageNum > 1)
                    {
                        await page.GoToAsync(url, new NavigationOptions
                        {
                            WaitUntil = new[] { WaitUntilNavigation.DOMContentLoaded },
                            Timeout = 30000
                        });
                    }

                    try
                    {
                        await page.WaitForSelectorAsync(".job-card.result", new WaitForSelectorOptions { Timeout = 10000 });
                    }
                    catch
                    {
                        _logger.LogWarning("No jobs found at page {pageNum}. Stopping.", pageNum);
                        break;
                    }

                    await Task.Delay(2000);

                    int jobCount = await page.EvaluateFunctionAsync<int>("() => document.querySelectorAll('.job-card.result').length");
                    _logger.LogInformation("Found {jobCount} jobs on page {pageNum}", jobCount, pageNum);

                    if (jobCount == 0)
                    {
                        _logger.LogWarning("No jobs found on page {pageNum}, stopping.", pageNum);
                        break;
                    }

                    int newJobsOnPage = 0;

                    for (int i = 0; i < jobCount; i++)
                    {
                        try
                        {
                            var jobInfo = await page.EvaluateFunctionAsync<JobInfo>($@"(index) => {{
                            const cards = Array.from(document.querySelectorAll('.job-card.result'));
                            const card = cards[index];
                            if (!card) return null;
                            
                            const titleLink = card.querySelector('.job-title a.job-link');
                            const company = card.querySelector('.job-company');
                            const location = card.querySelector('.job-location');
                            const badges = Array.from(card.querySelectorAll('.badge .content'));
                            const salary = badges.find(b => b.textContent.includes('$') || b.textContent.includes('hour') || b.textContent.includes('year'));
                            const posted = card.querySelector('.job-listed-date');
                            const jobAbstract = card.querySelector('.job-abstract');
                            
                            return {{
                                title: titleLink ? titleLink.textContent.trim() : '',
                                company: company ? company.textContent.trim() : '',
                                location: location ? location.textContent.trim() : '',
                                salary: salary ? salary.textContent.trim() : 'Not specified',
                                postedDate: posted ? posted.textContent.trim() : '',
                                url: titleLink ? titleLink.href : '',
                                shortDescription: jobAbstract ? jobAbstract.textContent.trim() : '',
                                description: '',
                                descriptionHtml: ''
                            }};
                        }}", i);

                            if (jobInfo == null || string.IsNullOrEmpty(jobInfo.Title))
                                continue;

                            if (processedUrls.Contains(jobInfo.Url))
                            {
                                _logger.LogDebug("Duplicate job skipped: {title}", jobInfo.Title);
                                continue;
                            }

                            processedUrls.Add(jobInfo.Url);

                            // Extract full job description
                            try
                            {
                                var clickSelector = $".job-card.result:nth-of-type({i + 1}) a.job-link.show-job-description";
                                await page.ClickAsync(clickSelector);
                                await page.WaitForSelectorAsync(".jdv-panel .job-description-container", new WaitForSelectorOptions { Timeout = 5000 });
                                await Task.Delay(2000);

                                var descData = await page.EvaluateFunctionAsync<DescriptionData>(@"() => {
                                const descContainer = document.querySelector('.jdv-panel .job-description-container');
                                if (!descContainer) return { text: '', html: '' };
                                return {
                                    text: (descContainer.innerText || '').trim(),
                                    html: (descContainer.innerHTML || '').trim()
                                };
                            }");

                                jobInfo.Description = CleanText(descData.Text);
                                jobInfo.DescriptionHtml = descData.Html;

                                _logger.LogDebug("Description extracted for {title} ({len} chars)", jobInfo.Title, jobInfo.Description.Length);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "Failed to get description for {title}", jobInfo.Title);
                                jobInfo.Description = jobInfo.ShortDescription;
                            }

                            jobs.Add(jobInfo);
                            newJobsOnPage++;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Error processing job index {index} on page {pageNum}", i, pageNum);
                        }
                    }

                    if (newJobsOnPage == 0)
                    {
                        _logger.LogWarning("No new jobs found on page {pageNum}, stopping.", pageNum);
                        break;
                    }

                    pageNum++;
                }

                _logger.LogInformation("Scraping complete. Total jobs collected: {count}", jobs.Count);
                await SaveJobsToExcel(jobs);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during scraping process.");
            }
        }

        private async Task<int> GetTotalPages(IPage page)
        {
            try
            {
                var totalPages = await page.EvaluateFunctionAsync<int>(@"() => {
                const indicator = document.querySelector('.search-results-page-number');
                if (!indicator) return 1;
                const text = indicator.textContent || '';
                const match = text.match(/of\s+(\d+)/i);
                return match && match[1] ? parseInt(match[1]) : 1;
            }");

                return totalPages > 0 ? totalPages : 1;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error detecting total pages. Defaulting to 1.");
                return 1;
            }
        }

        private string CleanText(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = Regex.Replace(text, @"[ \t]+", " ");
            text = Regex.Replace(text, @"\n\s*\n\s*\n+", "\n\n");
            return text.Trim();
        }

        public async Task SaveJobsToExcel(List<JobInfo> jobs, string fileName = "JobsFromJora.xlsx")
        {
            var exportDir = Path.Combine(Directory.GetCurrentDirectory(), "DataExports");
            if (!Directory.Exists(exportDir))
                Directory.CreateDirectory(exportDir);

            var filePath = Path.Combine(exportDir, fileName);

            if (File.Exists(filePath))
                File.Delete(filePath);

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Jobs");

            worksheet.Cells[1, 1].Value = "JobPostId";
            worksheet.Cells[1, 2].Value = "Title";
            worksheet.Cells[1, 3].Value = "Company";
            worksheet.Cells[1, 4].Value = "Location";
            worksheet.Cells[1, 5].Value = "Salary";
            worksheet.Cells[1, 6].Value = "Posted Date";
            worksheet.Cells[1, 7].Value = "URL";
            worksheet.Cells[1, 8].Value = "Description";

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

            _logger.LogInformation("Excel saved to: {filePath}", filePath);
        }
    }
}