namespace Personalized365.Web.ViewModels
{
    public class SummarizedMessage
    {
        public string Id { get; set; }

        public string Subject { get; set; }

        public DateTime ReceivedUtcDateTime { get; set; }

        public string BodyPreview { get; set; }

        public IList<string> SummarySentences { get; set; }
    }
}
