using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace MSGraphSendEmail.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmailController : ControllerBase
    {
        private readonly EmailUtility _emailUtility;
        public EmailController(EmailUtility emailUtility)
        {
            _emailUtility = emailUtility;
        }

        [HttpGet]
        public async Task<IActionResult> SendEmail()
        {
            string recipientEmail = "ToEmailAddress";
            await _emailUtility.SendMail(recipientEmail, string.Empty, string.Empty, "Test Email", "This is sample email", string.Empty);
            return Ok();
        }
    }
}
