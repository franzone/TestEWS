using Microsoft.Exchange.WebServices.Data;
using System;
using System.Linq;

namespace TestEWS
{
    class Program
    {
        public static string MY_EMAIL_ADDRESS = "";

        static void Main(string[] args)
        {
            try
            {
                // Connect to Exchange as a user logged into the same domain as the Exchange server
                ExchangeService service = new ExchangeService();
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(MY_EMAIL_ADDRESS, RedirectionUrlValidationCallback);
                Console.WriteLine("Autodiscovered URL : {0}", service.Url);

                // Open the INBOX
                Folder fInbox = Folder.Bind(service, WellKnownFolderName.Inbox);
                Console.WriteLine("Found {0} items in the INBOX", fInbox.TotalCount);

                // Find sub-folder for filtered emails
                Folder fFiltered = service.FindFolders(fInbox.Id, new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Filtered"), new FolderView(1)).FirstOrDefault();
                if (null != fFiltered)
                {
                    // Find email items at least 1 day old
                    TimeSpan tsOneDay = new TimeSpan(1, 0, 0, 0);
                    FindItemsResults<Item> oldItems = fFiltered.FindItems(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.Now.Subtract(tsOneDay)), new ItemView(1000));
                    Console.WriteLine("Found [{0}] old items", oldItems.TotalCount);

                    // Delete old items (put them into the Trash)
                    foreach (Item item in oldItems)
                    {
                        Console.WriteLine("Deleting : {0}", item.Subject);
                        item.Delete(DeleteMode.MoveToDeletedItems);
                    }
                }
                else
                {
                    Console.WriteLine("Filtered mailbox NOT FOUND");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error : {0}\r\n{1}", e.Message, e.StackTrace);
            }

            Console.Write("Press [Enter] to Continue...");
            Console.ReadLine();
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
