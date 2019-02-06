using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using StreamingNotificationsSample.Properties;

namespace StreamingNotificationsSample
{
	internal class Program
	{
		private static AutoResetEvent _Signal;
		private static ExchangeService _ExchangeService;
		private static string _SynchronizationState;
		private static Thread _BackroundSyncThread;

		private static StreamingSubscriptionConnection CreateStreamingSubscription(ExchangeService service,
		                                                                           StreamingSubscription subscription)
		{
			var connection = new StreamingSubscriptionConnection(service, 30);
			connection.AddSubscription(subscription);
			connection.OnNotificationEvent += OnNotificationEvent;
			connection.OnSubscriptionError += OnSubscriptionError;
			connection.OnDisconnect += OnDisconnect;
			connection.Open();

			return connection;
		}

		private static void SynchronizeChangesPeriodically()
		{
			while (true)
			{
				try
				{
					// Get all changes from the server and process them according to the business
					// rules.
					SynchronizeChanges(new FolderId(WellKnownFolderName.Calendar));
				}
				catch (Exception ex)
				{
					Console.WriteLine("Failed to synchronize items. Error: {0}", ex);
				}
				// Since the SyncFolderItems operation is a 
				// rather expensive operation, only do this every 10 minutes
				Thread.Sleep(TimeSpan.FromMinutes(10));
			}
		}

		public static void SynchronizeChanges(FolderId folderId)
		{
			bool moreChangesAvailable;
			do
			{
				Console.WriteLine("Synchronizing changes...");
				// Get all changes since the last call. The synchronization cookie is stored in the _SynchronizationState field.
				// Only the the ids are requested. Additional properties should be fetched via GetItem calls.
				var changes = _ExchangeService.SyncFolderItems(folderId, PropertySet.IdOnly, null, 512,
				                                               SyncFolderItemsScope.NormalItems, _SynchronizationState);
				// Update the synchronization cookie
				_SynchronizationState = changes.SyncState;

				// Process all changes
				foreach (var itemChange in changes)
				{
					// This example just prints the ChangeType and ItemId to the console
					// LOB application would apply business rules to each item.
					Console.Out.WriteLine("ChangeType = {0}", itemChange.ChangeType);
					Console.Out.WriteLine("ChangeType = {0}", itemChange.ItemId);
				}
				// If more changes are available, issue additional SyncFolderItems requests.
				moreChangesAvailable = changes.MoreChangesAvailable;
			} while (moreChangesAvailable);
		}


		public static void Main(string[] args)
		{
            ServicePointManager.ServerCertificateValidationCallback += sonicwallCertValidation;

            GetWebClientResponse();
            Console.Read();

            ////Utils.ShowTaskbarNotifier("testing testing");

            //// Create new exchange service binding
            //// Important point: Specify Exchange 2010 with SP1 as the requested version.
            //_ExchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
            //                   {
            //                       Credentials = new NetworkCredential(Settings.Default.User, Settings.Default.Password),
            //                       Url = new Uri(Settings.Default.ServerUrl)
            //                   };

            //// Process all items in the folder on a background-thread.
            //// A real-world LOB application would retrieve the last synchronization state first 
            //// and write it to the _SynchronizationState field.
            ////_BackroundSyncThread = new Thread(SynchronizeChangesPeriodically);
            ////_BackroundSyncThread.Start();

            //// Create a new subscription
            //var subscription = _ExchangeService.SubscribeToStreamingNotifications(new FolderId[] { WellKnownFolderName.Inbox },
            //                                                                      EventType.NewMail, EventType.Modified, EventType.Created, EventType.Moved);
            //// Create new streaming notification conection
            //var connection = CreateStreamingSubscription(_ExchangeService, subscription);

            //Console.Out.WriteLine("Subscription created.");
            //_Signal = new AutoResetEvent(false);

            //// Wait for the application to exit
            //_Signal.WaitOne();

            //// Finally, unsubscribe from the Exchange server
            //subscription.Unsubscribe();
            //// Close the connection
            //connection.Close();
        }

		private static void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
		{
			// Cast the sender as a StreamingSubscriptionConnection object.           
			var connection = (StreamingSubscriptionConnection) sender;
			// Ask the user if they want to reconnect or close the subscription. 
			Console.WriteLine("The connection has been aborted; probably because it timed out.");
			Console.WriteLine("Do you want to reconnect to the subscription? Y/N");
			while (true)
			{
				var keyInfo = Console.ReadKey(true);
				{
					switch (keyInfo.Key)
					{
						case ConsoleKey.Y:
							// Reconnect the connection
							connection.Open();
							Console.WriteLine("Connection has been reopened.");
							break;
						case ConsoleKey.N:
							// Signal the main thread to exit.
							Console.WriteLine("Terminating.");
							_Signal.Set();
							break;
					}
				}
			}
		}

		private static void OnNotificationEvent(object sender, NotificationEventArgs args)
		{
			// Extract the item ids for all NewMail Events in the list.
			var newMails = from e in args.Events.OfType<ItemEvent>()
			               where e.EventType == EventType.NewMail || e.EventType == EventType.Created || e.EventType == EventType.Modified || e.EventType == EventType.Deleted
			               select e.ItemId;

		    var ids = newMails as ItemId[] ?? newMails.ToArray();
		    if (!ids.Any())
		        return;

			// Note: For the sake of simplicity, error handling is ommited here. 
			// Just assume everything went fine
			var response = _ExchangeService.BindToItems(ids, new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Subject));
			var items = response.Select(itemResponse => itemResponse.Item);

			foreach (var item in items)
			{
                if (item == null) continue;
				//Console.Out.WriteLine("A new mail has been created. Received on {0}", item.DateTimeReceived);
				//Console.Out.WriteLine("Subject: {0}", item.Subject);
                Utils.ShowTaskbarNotifier(string.Format("Subject: {0}", item.Subject));
			}
		}

		private static void OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
		{
			// Handle error conditions. 
			var e = args.Exception;
			Console.Out.WriteLine("The following error occured:");
			Console.Out.WriteLine(e.ToString());
			Console.Out.WriteLine();
		}

        private static bool sonicwallCertValidation(object sender,
X509Certificate cert, X509Chain chain,
System.Net.Security.SslPolicyErrors error)
        {
            return true;
        }

	    // store cookies so we don't always need to authenticate
        static readonly CookieContainer cookieJar = new CookieContainer();
	    private const string url = "https://192.168.192.201/ws/reports";
	    private static WebResponse response;
	    public static string GetWebClientResponse()
        {
            var document = "";
            try
            {
                var password = performPasswordMD5HashCompute("DMacLeod1592");

                var request = WebRequest.Create(url) as HttpWebRequest;
                if (request != null)
                {
                    request.Credentials = new NetworkCredential("admin", password);
                    // save the cookie
                    request.CookieContainer = cookieJar;

                    response = request.GetResponse();
                    var reader = new StreamReader(response.GetResponseStream());
                    document = reader.ReadToEnd();
                    
                    response.Close();
                    Debug.WriteLine("SonicWall Done");
                }
            }
            catch (WebException we)
            {
                //log the error
                Debug.WriteLine(we.Message);
                Debug.WriteLine("An error has ocurred in Sonicwall Web Service. Error Detail:\n" + we.Message);
                Debug.WriteLine("An error occurred while processing Sonicwall URI: " + url);
            }
            return document;
        }

        private static string performPasswordMD5HashCompute(string password)
        {
            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                Byte[] originalBytes = Encoding.Default.GetBytes(password);
                Byte[] encodedBytes = md5.ComputeHash(originalBytes);
                //return BitConverter.ToString(encodedBytes).ToLower();//.Replace("-", "");
                return Convert.ToBase64String(encodedBytes).ToLower();
            }
        }
	}
}