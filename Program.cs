using System;
using Microsoft.Exchange.WebServices.Data;

namespace v20201027ews
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            // dotnet add package Microsoft.Exchange.WebServices --version 2.2.0
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications?redirectedfrom=MSDN
            string fromEmail = "user1@contoso.com";
            string fromPwd = "password";
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(fromEmail, fromPwd);
            service.UseDefaultCredentials = true;
            service.AutodiscoverUrl(fromEmail, RedirectionUrlValidationCallback);

            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("abc@exammple.org");
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API.");
            email.Send();

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
