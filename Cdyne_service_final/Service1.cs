/* http://messaging.cdyne.com/Messaging.svc?wsdl was added as Service Reference and given the name WSDL */

using System;
using System.ServiceProcess;
using Cdyne_service_final.WSDL;
using Microsoft.Exchange.WebServices.Data;
using System.Diagnostics;

namespace Cdyne_service_final
{
    public partial class CDYNESMS : ServiceBase
    {

        private System.Timers.Timer timer;
        EventLog appLog = new EventLog();

        public CDYNESMS()
        {
            this.ServiceName = "Cdyne Service";
            InitializeComponent();
            
        }

        static void Main(string[] args)
        {
            Run(new CDYNESMS());
        }

        protected override void OnStart(string[] args)
        {
            this.timer = new System.Timers.Timer(30000D);  // 30000 milliseconds = 30 seconds
            this.timer.AutoReset = true;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
            this.timer.Start();

            appLog.Source = "Application";
            appLog.WriteEntry("CDYNE SMS Service Started", EventLogEntryType.Information);

        }

        protected override void OnStop()
        {
            this.timer.Stop();
            this.timer = null;

            appLog.Source = "Application";
            appLog.WriteEntry("CDYNE SMS Service Stopped", EventLogEntryType.Information);
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                cdynesvc();
            }
            catch (Exception)
            {
                appLog.Source = "Application";
                appLog.WriteEntry("CDYNE SMS Service Stopped Unexpectedly", EventLogEntryType.Error);
            }
        }

        static void cdynesvc()
        {


            EventLog appLog = new EventLog();
            appLog.Source = "Application";

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            string tophone;

            MessagingClient client = new MessagingClient("mms2wsHttpBinding");
            OutgoingMessageRequest req = new OutgoingMessageRequest();
            req.LicenseKey = new Guid(""); //put your CDYNE API Key here
            req.UseMMS = false;

            service.Credentials = new WebCredentials("username", "password", "domainname.local");
            service.Url = new Uri("https://mail.contoso.com/EWS/Exchange.asmx");  //put the URL to your Exchange EWS server here.
            service.UseDefaultCredentials = false;

            FindItemsResults<Item> findResults = service.FindItems(
            WellKnownFolderName.Inbox,
            new ItemView(1));

            foreach (Item item in findResults.Items)
            {
                if (item.Subject.Contains("Keyword"))
                {
                    tophone = "5555555555"; //cell phone number in the US
                    req.To = new string[] { tophone };

                    req.Body = item.Subject;
                    OutgoingMessageResponse[] resp = client.SendMessage(req);
                    Console.Write(item.Subject);

                    item.Delete(DeleteMode.HardDelete);

                    appLog.WriteEntry("CDYNE SMS Sent Keyword SMS", EventLogEntryType.Information);

                }
                else
                {
                    tophone = "5555555556";
                    req.To = new string[] { tophone };

                    req.Body = item.Subject;
                    OutgoingMessageResponse[] resp = client.SendMessage(req);
                    Console.Write(item.Subject);

                    item.Delete(DeleteMode.HardDelete);

                    appLog.WriteEntry("CDYNE SMS Sent Other SMS", EventLogEntryType.Information);

                }
            }
        }
    }
}
