namespace OutlookSpamDetectionWeb.Models
{
    public class EmailInfo
    {
        public Email From { get; set; }
        public Email To { get; set; }
        public string Subject { get; set; }
        public string BodyText { get; set; }
    }
}