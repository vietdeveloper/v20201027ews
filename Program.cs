using System;
using Microsoft.Exchange.WebServices.Data;

namespace v20201027ews
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("{0} Hello World!", DateTime.Now);

            // dotnet add package Microsoft.Exchange.WebServices --version 2.2.0
            // dotnet remove package Microsoft.Exchange.WebServices
            // dotnet add package Exchange.WebServices.Managed.Api
            // dotnet remove package Exchange.WebServices.Managed.Api
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications?redirectedfrom=MSDN
            string fromEmail = "user1@contoso.com";
            string fromPwd = "password";

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(fromEmail, fromPwd);
            //If your client targets an Exchange Online or Office 365 Developer Site mailbox, verify that UseDefaultCredentials is set to false, which is the default value. 
            //Comment out to make it work.            
            //service.UseDefaultCredentials = true;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.AutodiscoverUrl(fromEmail, RedirectionUrlValidationCallback);

            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("abc@example.org"); // Remote Server returned '550 5.1.3 STOREDRV.Submit; invalid recipient address'
            email.Subject = "Hello World from .NET and Visual Studio Code";
            email.Body = new MessageBody("This is the test email I've sent by using the EWS Managed API. Cannot use this library with .NET Core 3.1 framework.");
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
