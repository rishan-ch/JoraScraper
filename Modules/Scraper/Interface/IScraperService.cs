using JoraScraper.Modules.Scraper.Dto;

namespace JoraScraper.Modules.Scraper.Interface
{
    public interface IScraperService
    {
        Task SaveJobsToExcel(List<JobInfo> jobs, string fileName = "JobsFromJora.xlsx");
        Task ScrapeAndSaveJobsAsync();
    }
}
