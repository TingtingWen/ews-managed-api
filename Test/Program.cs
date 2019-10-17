using Microsoft.Exchange.WebServices.Data;
using System;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2016);
            //service.Credentials = new WebCredentials("MSOXWSCORE_User01@contoso.com", "Ap6zF1#uP4l-+");
            service.Credentials = new WebCredentials("plugdevuser01@contoso.com", "F@2019ang");
            service.UseDefaultCredentials = true;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            //service.AutodiscoverUrl("MSOXWSCORE_User01@contoso.com", RedirectionUrlValidationCallback);
            service.Url = new Uri("http://sut01.contoso.com/EWS/Exchange.asmx");


            //EmailMessage email = new EmailMessage(service);
            //email.ToRecipients.Add("MSOXWSCORE_User01@contoso.com");
            //email.Subject = "HelloWorld";
            //email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            //email.Send();

            //ChangeCollection<ItemChange> result = service.SyncFolderItems(WellKnownFolderName.Inbox, PropertySet.IdOnly, null, 100, SyncFolderItemsScope.NormalItems, null);

            //List<ItemId> ids = new List<ItemId>();
            //ids.Add(result[0].ItemId);

            //service.ArchiveItems(ids, WellKnownFolderName.Inbox);

            #region MS-OXWSEDISC.opn
            // Add the related user to admin role "Discovery Management" (Assigned Roles: "Legal Hold" and "Mailbox Search") in ECA.
            SetHoldOnMailboxesParameters parameters = new SetHoldOnMailboxesParameters();
            parameters.ActionType = HoldAction.Create;
            parameters.HoldId = "HoldId2";
            parameters.Query = "test";
            parameters.Mailboxes = new string[1] { "/o=CONTOSO/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=f1983acba6364b1abab45687feedb253-MSOXWSCORE_User0" };
            service.SetHoldOnMailboxes(parameters);

            System.Threading.Thread.Sleep(1000);
            service.GetHoldOnMailboxes("HoldId2");

            System.Threading.Thread.Sleep(1000);
            string[] userRoles = new string[2] { "MailboxSearch", "LegalHold" };
            ManagementRoles managementRoles = new ManagementRoles(userRoles, null);
            service.ManagementRoles = managementRoles;
            service.GetSearchableMailboxes(string.Empty, false);

            System.Threading.Thread.Sleep(1000);
            SearchMailboxesParameters searchMailboxesParameters = new SearchMailboxesParameters();
            searchMailboxesParameters.ResultType = SearchResultType.StatisticsOnly;
            searchMailboxesParameters.SearchQueries = new MailboxQuery[1];
            MailboxSearchScope mailboxSearchScope = new MailboxSearchScope(
                "/o=CONTOSO/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=f1983acba6364b1abab45687feedb253-MSOXWSCORE_User0",
                MailboxSearchLocation.All);
            searchMailboxesParameters.SearchQueries[0] = new MailboxQuery("Test", new MailboxSearchScope[1] { mailboxSearchScope });
            service.SearchMailboxes(searchMailboxesParameters);

            System.Threading.Thread.Sleep(1000);
            service.GetDiscoverySearchConfiguration("HoldId2", false, false);
            #endregion

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return true;
            //// The default for the validation callback is to reject the URL.
            //bool result = false;
            //Uri redirectionUri = new Uri(redirectionUrl);
            //// Validate the contents of the redirection URL. In this simple validation
            //// callback, the redirection URL is considered valid if it is using HTTPS
            //// to encrypt the authentication credentials. 
            //if (redirectionUri.Scheme == "https")
            //{
            //    result = true;
            //}
            //return result;
        }
    }
}
