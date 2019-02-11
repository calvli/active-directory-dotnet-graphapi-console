namespace GraphConsoleAppV3
{
    internal class AppModeConstants
    {
        public const string ClientSecret = "";
        public const string TenantName = "";
        public const string AuthString = GlobalConstants.AuthString + TenantName;
    }

    internal class UserModeConstants
    {
        public const string AuthString = GlobalConstants.AuthString + "common/";
    }

    internal class GlobalConstants
    {
        public const string AuthString = "https://login.microsoftonline.com/";
        public const string ResourceUrl = "https://graph.windows.net/";
        public const string GraphServiceObjectId = "00000002-0000-0000-c000-000000000000";
        public const string TenantId = "72f988bf-86f1-41af-91ab-2d7cd011db47";
        public const string ClientId = "dacf386f-a055-4a78-8ade-6365265f020b";
    }
}
