using MailKit.Net.Imap;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Downloader
{
    class Program
    {
        private static int deleteOverXDays = 30;
        static void Main(string[] args)
        {
            if (args.Length >= 4)
            {
                if (!System.IO.Directory.Exists("Download"))
                {
                    System.IO.Directory.CreateDirectory("Download");
                }

                int.TryParse(ConfigurationManager.AppSettings["DeleteEmailOverDaysOld"], out deleteOverXDays);

                if (args.Length == 6)
                {
                    int.TryParse(args[5], out deleteOverXDays);
                }

                if (deleteOverXDays <= 0)
                {
                    deleteOverXDays = 31;
                }

                string userName = args[2];
                string password = args[3];

                bool useSSL = false;
                if (args.Length == 5)
                {
                    useSSL = args[4] == "1";
                }

                ImapClient imap = new ImapClient();
                if (useSSL)
                {
                    imap.Connect(args[0], Convert.ToInt32(args[1]), MailKit.Security.SecureSocketOptions.SslOnConnect);
                }
                else
                {
                    imap.Connect(args[0], Convert.ToInt32(args[1]), false);
                }
                imap.Authenticate(userName, password);
                
                Console.WriteLine("Logged into the IMAP server");

                // Get a list of the folders and walk them
                foreach (var name in imap.GetFolders(imap.PersonalNamespaces[0]))
                {
                    ProcessFolder(ref imap, name);
                }

                imap.Dispose();
            }
            else
            {
                Console.WriteLine("Syntax:");
                Console.WriteLine("Downloader <server> <port> <user> <pass> <usessl> <deleteoverxdays>");
                Console.WriteLine("server, port, user and pass are required");
                Console.WriteLine("usessl: can be 0 => Disable, 1 => Enabled");
                Console.WriteLine("deleveoverxdays: (Optional) If provided will override the value in config file");
            }
        }

        private static bool IgnoreThisFolder(string folderName)
        {
            string conf = ConfigurationManager.AppSettings["IgnoreFolders"];
            bool ignoreInboxRoot = bool.Parse(ConfigurationManager.AppSettings["IgnoreInbox"]);
            string[] folders = conf.Split(';');

            if (ignoreInboxRoot && folderName.ToLower() == "inbox")
            {
                return true;
            }

            foreach (string folder in folders)
            {
                if (folderName.ToLower().StartsWith(folder))
                {
                    Console.WriteLine("Ignoring folder {0}", folderName);
                    return true;
                }
            }
            return false;
        }

        private static void ProcessFolder(ref ImapClient imap, MailKit.IMailFolder folder)
        {
            if (IgnoreThisFolder(folder.FullName))
            {
                return;
            }

            string folderPath = folder.FullName.Replace(".", "\\").Replace("&", " and ");
            Console.WriteLine("{0} => {1}", folder, folderPath);

            if (!System.IO.Directory.Exists("Download\\" + folderPath))
            {
                System.IO.Directory.CreateDirectory("Download\\" + folderPath);
            }
            folder.Open(MailKit.FolderAccess.ReadWrite);
            for (int x = 0; x < folder.Count; x++)
            {
                Console.WriteLine("Checking message {0}", x);
                var summary = folder.Fetch(new List<int>() { x }, MailKit.MessageSummaryItems.All)[0];
                if (!summary.Flags.Value.HasFlag(MailKit.MessageFlags.Deleted))
                {
                    Console.WriteLine("Downloading");
                    try
                    {
                        // Download the message if we have not already
                        var str = folder.GetStream(x, string.Empty);

                        string fileName = Sanitise(((MimeKit.MailboxAddress)summary.Envelope.From[0]).Address.ToString()) + "_" + summary.Envelope.Date.Value.ToString("ddMMyyyy_hhmmss") + "_" + Sanitise(summary.Envelope.MessageId) + ".eml";

                        using (System.IO.StreamWriter file =
                            new System.IO.StreamWriter(string.Format("Download\\{1}\\{0}", fileName, folderPath)))
                        {
                            str.CopyTo(file.BaseStream);
                        }

                        if (summary.Envelope.Date.Value <= DateTime.UtcNow.AddDays(deleteOverXDays * -1))
                        {
                            Console.WriteLine("Removing old message from this folder");
                            folder.AddFlags(x, MailKit.MessageFlags.Deleted, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(" ==> Did not download due to error {0}", ex.Message);
                    }
                }
            }
            folder.Close();
        }

        private static string Sanitise(string input)
        {
            if (string.IsNullOrEmpty(input))
            { return string.Empty; }

            return input.Replace("@", "").Replace(".", "").Replace("<", "").Replace(">", "").Replace(" ", "").Replace("!", "").Replace("\\", "").Replace("/", "");
        }
    }
}
