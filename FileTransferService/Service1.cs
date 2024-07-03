using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace FileTransferService.NET
{
    public partial class Service1 : ServiceBase
    {
        Timer Timer = new Timer();
        int Interval = 10000; // 10000 ms = 10 second  

        public Service1()
        {
            InitializeComponent();
            this.ServiceName = "FileTransferService.NET";
        }

        protected override void OnStart(string[] args)
        {
            WriteLog("Service is started at " + DateTime.Now);
            Timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            Timer.Interval = Interval;
            Timer.Enabled = true;

            //Connect to NDM Share
            NetworkCredential credentials = new NetworkCredential();
            using (new ConnectToSharedFolder("networkPath", credentials))
            {
                StreamReader str = new StreamReader(@"\\192.168.0.1\C$\Test\test.txt");
                var x = str.ReadToEnd();
            }
            //Upload to Sharepoint
            string userName = "xxx@xxxx.onmicrosoft.com";
            string password = "xxx";
            Uploadfiles(userName, password);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteLog("Service is recall at " + DateTime.Now);
        }

        protected override void OnStop()
        {
            Timer.Stop();
            WriteLog("Service is stopped at " + DateTime.Now);
        }

        public void WriteLog(string logMessage, bool addTimeStamp = true)
        {
            var path = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            var filePath = String.Format("{0}\\{1}_{2}.txt",
                path,
                ServiceName,
                DateTime.Now.ToString("yyyyMMdd", CultureInfo.CurrentCulture)
                );

            if (addTimeStamp)
                logMessage = String.Format("[{0}] - {1}\r\n",
                    DateTime.Now.ToString("HH:mm:ss", CultureInfo.CurrentCulture),
                    logMessage);

            System.IO.File.AppendAllText(filePath, logMessage);
        }

        static string siteUrl = "https://sppalsmvp.sharepoint.com/sites/DeveloperSite/";

        public async static void Uploadfiles(string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            using (var clientContext = new ClientContext("https://xxx.sharepoint.com/sites/test"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                Web web = clientContext.Web;
                clientContext.Load(web, a => a.ServerRelativeUrl);
                clientContext.ExecuteQuery();
                List documentsList = clientContext.Web.Lists.GetByTitle("Contact");

                var fileCreationInformation = new FileCreationInformation();
                //Assign to content byte[] i.e. documentStream  

                fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"D:\document.pdf");
                //Allow owerwrite of document  

                fileCreationInformation.Overwrite = true;
                //Upload URL  

                fileCreationInformation.Url = "https://testlz.sharepoint.com/sites/jerrydev/" + "Contact/demo" + "/document.pdf";

                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);

                //Update the metadata for a field having name "DocType"  
                uploadFile.ListItemAllFields["Title"] = "UploadedviaCSOM";

                uploadFile.ListItemAllFields.Update();
                clientContext.ExecuteQuery();

            }
        }
    }
}
