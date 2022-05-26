using System;

namespace MSGraphSendEmail
{
    public class AccessToken
    {        
        public AccessToken(
            string token,
            DateTime expiresOn)
        {
            Token = token;
            Expires = expiresOn;
        }
     
        public string Token { get; set; }

        public DateTime Expires { get; }        

        public bool Expired => DateTime.UtcNow > Expires;
    }
}
