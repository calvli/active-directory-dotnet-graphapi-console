#region

using System;
using System.CodeDom.Compiler;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Data;
using System.Net.WebSockets;
using System.Text;

#endregion

namespace GraphConsoleAppV3
{
    public class Program
    {

        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            Console.WriteLine("Run operations for signed-in user, or in app-only mode.\n");
            Console.WriteLine("[a] - app-only\n[u] - as user\n[b] - both as user first, and then as app.\nPlease enter your choice:\n");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    Console.WriteLine("\nRunning app-only mode\n\n");
                    var appModeTask = Task.Run(() => Requests.AppMode());
                    appModeTask.Wait();
                    break;
                case 'b':
                    Console.WriteLine("\nRunning app-only mode, followed by user mode\n\n");
                    appModeTask = Task.Run(() => Requests.AppMode());
                    appModeTask.Wait();
                    var userModeTask = Task.Run(() => Requests.UserMode());
                    userModeTask.Wait();
                    break;
                case 'u':
                    Console.WriteLine("\nRunning in user mode\n\n");
                    userModeTask = Task.Run(() => Requests.UserMode());
                    userModeTask.Wait();
                    break;
                default:
                    WriteError("\nSelection not recognized. Running in user mode\n\n");
                    userModeTask = Task.Run(() => Requests.UserMode());
                    userModeTask.Wait();
                    break;
            }

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", currentDateTime);
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
