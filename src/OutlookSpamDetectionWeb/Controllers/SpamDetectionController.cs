using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OutlookSpamDetectionWeb.Models;

namespace OutlookSpamDetectionWeb.Controllers
{
    [ApiController]
    [Route("api/email")]
    public class OutlookSpamDetection : ControllerBase
    {
        private readonly ILogger<OutlookSpamDetection> _logger;

        public OutlookSpamDetection(ILogger<OutlookSpamDetection> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public EmailResponse CheckEmail([FromBody] EmailInfo emailInfo)
        {
            var message = emailInfo.BodyText.Replace("\n", string.Empty);

            return new EmailResponse { IsSpam = true };
        }
    }
}