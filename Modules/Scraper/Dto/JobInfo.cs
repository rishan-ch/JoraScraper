namespace JoraScraper.Modules.Scraper.Dto
{
    public class JobInfo
    {
        public Guid JobPostId { get; set; } = Guid.NewGuid();
        public string? Title { get; set; }
        public string? Company { get; set; }
        public string? Location { get; set; }
        public string? Salary { get; set; }
        public string? PostedDate { get; set; }
        public string? Url { get; set; }
        public string? ShortDescription { get; set; }
        public string? Description { get; set; }
        public string? DescriptionHtml { get; set; }
    }

    public class DescriptionData
    {
        public string? Text { get; set; }
        public string? Html { get; set; }
    }
}
