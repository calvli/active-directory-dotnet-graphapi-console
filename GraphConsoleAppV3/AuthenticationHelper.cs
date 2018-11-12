using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace GraphConsoleAppV3
{
    internal class AuthenticationHelper
    {
        public static string TokenForUser;
        public static string TokenForApplication;

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsApplication()
        {
            Uri servicePointUri = new Uri(GlobalConstants.ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, GlobalConstants.TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForApplication());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication()
        {
            return await GetTokenForApplication();
        }

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static async Task<string> GetTokenForApplication()
        {
            if (TokenForApplication == null)
            {
                AuthenticationContext authenticationContext = new AuthenticationContext(
                    AppModeConstants.AuthString,
                    false);

                // Configuration for OAuth client credentials 
                if (string.IsNullOrEmpty(AppModeConstants.ClientSecret))
                {
                    Program.WriteError(
                        "Client secret not set. Please follow the steps in the README to generate a client secret.");
                }
                else
                {
                    ClientCredential clientCred = new ClientCredential(
                        GlobalConstants.ClientId,
                        AppModeConstants.ClientSecret);
                    AuthenticationResult authenticationResult =
                        await authenticationContext.AcquireTokenAsync(GlobalConstants.ResourceUrl, clientCred);
                    TokenForApplication = authenticationResult.AccessToken;
                }
            }
            return TokenForApplication;
        }

        /// <summary>
        /// Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        {
            Uri servicePointUri = new Uri(GlobalConstants.ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, GlobalConstants.TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForUser());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> AcquireTokenAsyncForUser()
        {
            return await GetTokenForUser();
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                TokenForUser = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IndVTG1ZZnNxZFF1V3RWXy1oeFZ0REpKWk00USIsImtpZCI6IndVTG1ZZnNxZFF1V3RWXy1oeFZ0REpKWk00USJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNTI0ODdkODAtMTMzZi00OTcwLTgwNzEtNWI0YjUyZDI5M2UzLyIsImlhdCI6MTU0MTgwMTY4NywibmJmIjoxNTQxODAxNjg3LCJleHAiOjE1NDE4MDU1ODcsImNsb3VkX2luc3RhbmNlX25hbWUiOiJtaWNyb3NvZnRvbmxpbmUudXMiLCJjbG91ZF9ncmFwaF9ob3N0X25hbWUiOiJncmFwaC53aW5kb3dzLm5ldCIsIm1zZ3JhcGhfaG9zdCI6ImdyYXBoLm1pY3Jvc29mdC5jb20iLCJhY3IiOiIxIiwiYWlvIjoiWTJCZ1lMaDZNTWhWS21FR3c0TlpEczFzeHh5Vy8rQVByWFZrQ0JTeDB2aVNOTWMyNnpJQSIsImFtciI6InB3ZCIsImFwcGlkIjoiZDNjZTRjZjgtNjgxMC00NDJkLWI0MmUtMzc1ZTE0NzEwMDk1IiwiYXBwaWRhY3IiOiIwIiwiZV9leHAiOiIyNjI4MDAiLCJpcGFkZHIiOiIxMzEuMTA3LjE2MC4xNzciLCJuYW1lIjoiYWRtaW43Iiwib2lkIjoiZWM0YjBkNDctMjkyNi00YzNiLTgyOGEtN2FiMWZmYjIwM2E3IiwicHVpZCI6IjEwMDMyOTE4NTQ2QjYxNTgiLCJzY3AiOiJEaXJlY3RvcnkuQWNjZXNzQXNVc2VyLkFsbCBVc2VyLlJlYWQiLCJzdWIiOiJXZ2ZTcEJWSE1yWkFnY25vaEJRSHNlY2ZaOHgyb0ppcXlWY1kzZHJvYTU0IiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IlVTR292IiwidGlkIjoiNTI0ODdkODAtMTMzZi00OTcwLTgwNzEtNWI0YjUyZDI5M2UzIiwidW5pcXVlX25hbWUiOiJhZG1pbjdAQXJsRTJFLm9ubWljcm9zb2Z0LnVzIiwidXBuIjoiYWRtaW43QEFybEUyRS5vbm1pY3Jvc29mdC51cyIsInV0aSI6IlFCbmNIbEkzdDBlcWQ5NW0tM1lNQUEiLCJ2ZXIiOiIxLjAifQ.I9AqaCBl_qlR726_1XoAGlxS623KMFkRHueM2-i14jnEkkj_yPyr23SXxO_CzXS9DykmIoHhd9wJiCMoa9CK7xMN-rktWhzVNuw9egN6JI7WJDnpPDG-YXtuFTAcAgbwXUaY5JDT2nXEUystX8SdR_rt77COE1fYEtHeuIG9BqPTC8zS_xD_8TOiHw9GhOlu7A-Gg6iDCerjop3YbN_iOIaKMtenmZK8ijKdtjE8A8T4nUwuYwv_Y-uaAjHGUjWTn44U-vnP8gqcdP5aoO79dZ9IogxdyIiRwxYEYbctKB68KXIeyIhJqtY60FZbm-dVUOyyZw61v4f46rO9q8oAVg";
                //var redirectUri = new Uri("https://localhost");
                //AuthenticationContext authenticationContext = new AuthenticationContext(UserModeConstants.AuthString, false);
                //AuthenticationResult userAuthnResult = await authenticationContext.AcquireTokenAsync(GlobalConstants.ResourceUrl,
                //    GlobalConstants.ClientId, redirectUri, new PlatformParameters(PromptBehavior.RefreshSession));
                //TokenForUser = userAuthnResult.AccessToken;
                //Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                //                  userAuthnResult.UserInfo.FamilyName);
            }
            return TokenForUser;
        }

    }
}
