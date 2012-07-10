
namespace crmail
{
    using System;
    using System.Globalization;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using System.Configuration;

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0 || args.Length > 1)
            {
                ShowSyntax();
                return;
            }

            try
            {
                ReadConfig();

                // Connect to the server
                Console.Write(string.Format(CultureInfo.InvariantCulture, Resources.InfoConnecting,
                            Program.ServerName));

                TeamFoundationServer tfServer = TfServerUtil.GetServer(Program.ServerName);
                Console.WriteLine(Resources.InfoConnected);

                VersionControlServer sccServer = 
                    (VersionControlServer)tfServer.GetService(typeof(VersionControlServer));

                // call the email handler to go create the email.
                MailFromShelveset mailFromShelveset = new MailFromShelveset(sccServer, args[0], CrAlias, WebViewPort);
                mailFromShelveset.GenerateMail();
            }
            // Catch all, generally bad but fine here as we are just dumping error and closing
            catch (Exception exception)
            {
                // put some fancy red color
                ConsoleColor initColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine(string.Format(CultureInfo.InvariantCulture, 
                                   Resources.ErrFail, exception.Message));

                // remember to reset to original color
                Console.ForegroundColor = initColor;
            }
        }

        /// <summary>
        /// Read app.config
        /// </summary>
        private static void ReadConfig()
        {
            ServerName = ConfigurationManager.AppSettings["ServerName"];
            CrAlias = ConfigurationManager.AppSettings["CrAlias"];
            string webViewPortStr = ConfigurationManager.AppSettings["WebViewPort"];
            if (!string.IsNullOrEmpty(webViewPortStr))
            {
                int port;
                if (int.TryParse(webViewPortStr, out port))
                {
                    WebViewPort = port;
                }
            }

            if (string.IsNullOrEmpty(ServerName) || WebViewPort == 0)
            {
                throw new ApplicationException(Resources.ErrBadConfigFile);
            }
        }

        /// <summary>
        /// Show usage syntax.
        /// </summary>
        private static void ShowSyntax()
        {
            Console.WriteLine(Resources.InfoHelpDescription);
            Console.WriteLine();
            Console.WriteLine("crmail shelvesetname"); // no loc needed for this
            Console.WriteLine();
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, Resources.InfoHelpSample, 
                @"crmail myshelveset;mydomain\myalias"));
        }

        /// <summary>
        /// Name of the TFS server being used.
        /// </summary>
        public static string ServerName
        {
            get;
            set;
        }

        /// <summary>
        /// Email alias of the code review distribution list
        /// </summary>
        public static string CrAlias
        {
            get;
            set;
        }

        public static int WebViewPort
        {
            get;
            set;
        }
    }
}
