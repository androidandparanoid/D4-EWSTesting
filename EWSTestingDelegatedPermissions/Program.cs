using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;

namespace EwsOAuth
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).Wait();

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            }
        }

        static async System.Threading.Tasks.Task MainAsync(string[] args)
        {
            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ConfigurationManager.AppSettings["appId"],
                TenantId = ConfigurationManager.AppSettings["tenantId"]
            };

            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            var ewsScopes = new string[] { "https://outlook.office.com/EWS.AccessAsUser.All" };

            try
            {
                // Make the interactive token request
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

                // Make an EWS call
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                foreach (var folder in folders)
                {
                    Console.WriteLine($"Folder: {folder.DisplayName}");
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex.ToString()}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }
        }
    }
}