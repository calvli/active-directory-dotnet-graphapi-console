#region

using System;
using System.Net;

#endregion

namespace GraphConsoleAppV3
{
    public class Program
    {

        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            Console.WriteLine("Run operations for signed-in user, or in app-only mode.\n");
            Console.WriteLine("[a] - app-only\n[u] - as user\n[b] - both as user first, and then as app.\nPlease enter your choice:\n");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    Console.WriteLine("\nRunning app-only mode\n\n");
                    Requests.AppMode().Wait();
                    break;
                case 'b':
                    Console.WriteLine("\nRunning app-only mode, followed by user mode\n\n");
                    Requests.AppMode().Wait();
                    Requests.UserMode().Wait();
                    break;
                case 'u':
                    Console.WriteLine("\nRunning in user mode\n\n");
                    Requests.UserMode().Wait();
                    break;
                default:
                    WriteError("\nSelection not recognized. Running in user mode\n\n");
                    Requests.UserMode().Wait();
                    break;
            }

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", DateTime.Now.ToUniversalTime());
            Console.ReadKey();
        }

        public static void WriteError(string output, params object[] args)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine(output, args);
            Console.ResetColor();
        }
    }
}
