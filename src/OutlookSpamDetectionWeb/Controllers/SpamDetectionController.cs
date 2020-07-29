using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OutlookSpamDetectionML.Model;
using OutlookSpamDetectionWeb.Models;

namespace OutlookSpamDetectionWeb.Controllers
{
    [ApiController]
    [Route("api/email")]
    public class SpamDetection : ControllerBase
    {
        private readonly ILogger<SpamDetection> _logger;

        public SpamDetection(ILogger<SpamDetection> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public EmailResponse CheckEmail([FromBody] EmailInfo emailInfo)
        {
            var message = emailInfo.BodyText.Replace("\n", string.Empty);
            var model = ConsumeModel.Predict(new ModelInput { Message = message });

            return new EmailResponse { IsSpam = model.Prediction };
        }
    }
}