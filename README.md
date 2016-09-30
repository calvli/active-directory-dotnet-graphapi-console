---
services: active-directory
platforms: dotnet
author: dstrockis
---

# Call the Azure AD Graph API from a native client

Console App using Graph Client Library Version 2.0

This console application is a .Net sample, using the Graph API Client library (Version 2.0) - it demonstrates common Read calls to the Graph API including Getting Users, Groups, Group Membership, Roles, Tenant information, Service Principals, Applications. The second part of the sample app demonstrates common Write/Update/Delete options on Users, Groups, and shows how to execute User License Assignment, updating a User's thumbnailPhoto and links, etc.  It also can read the contents from the signed-on user's mailbox.

The sample incorporates using the Active Directory Authentication Library (ADAL) for authentication. The first part of the console app is Read-only, and uses OAuth Client Credentials to authenticate against the Demo Company. The second part of the app has update operations, and requires User credentials (using the OAuth Authorization Code flow). Update operations are not permitted using this shared Demo company - to try these, you will need to update the application configuration to be used with your own Azure AD tenant, and will need to configure applications accordingly. When configuring this application to be used with your own tenant, you will need to configure two applications in the Azure Management Portal: one using OAuth Client Credentials, and a second one using OAuth Authorization Code flow (each with separate ClientIds (AppIds)) - to execute update operations, you will need to logon with an account that has Administrative permissions.  This app also demonstrates how to read the mailbox of the signed on user account using the Microsoft common consent framework  - applications can be granted access to Azure Active directory, as well as Microsoft Office365 applications including Exchange and SharePoint Online.  To configure the app:


Step 1: Clone or download this repository
From your shell or command line:

`git clone https://github.com/Azure-Samples/active-directory-dotnet-graphapi-console.git`


Step 2: Run the sample in Visual Studio 2013
The sample app is preconfigured to read data from a Demonstration company (GraphDir1.onMicrosoft.com) in Azure AD. 
Run the sample application by selecting F5.  The second part of the app will require Admin credentials, you can simulate 
authentication using this demo user account: userName =  demoUser@graphDir1.onMicrosoft.com, password = graphDem0 
 However, this is only a user account and does not have administrative permissions to execute updates - therefore, you
will see "..unauthorized.." response errors when attempting any requests requiring admin permissions.  To see how updates
work, you will need to configure and use this sample with your own tenant - see the next step.


Step 3: Running this application with your own Azure Active Directory tenant

Register the Sample app for your own tenant

1. Sign in to the Azure management portal at [http://wwww.windowsazure.com](http://www.windowsazure.com).

2. Type in **App registrations** in the search bar.

3. Enter a friendly name for the application, for example **"Console App for Azure AD"**, select **"Web Application and/or Web API"**, and click **next**. 

4. For the **Sign-on URL**, enter a value (NOTE: this is not used for the console app, so is only needed for this initial configuration):  "http://localhost"

5. For the **App ID URI**, enter "http://localhost".  Click the checkmark to complete the initial configuration.

6. Click **create**.

7. Find the **Application ID** value and copy it aside, you will need this later when configuring your application.

8. Under the Keys section, select either a 1-year or 2-year key - the **keyValue** will be displayed after you save the configuration at the end - it will be displayed, and you should save this to a secure location. NOTE: The key value is only displayed once, and you will not be able to retrieve it later.

9. Configure Permissions - under the **"Required permissions"** section, you will configure permissions to access the Graph (Azure Active Directory). Select "Windows Azure Active Directory", then select "Read directory data". Notes: this configures the App to use OAuth Client Credentials, and have Read access permissions for the application. 

10. Click the **Save** button near the top of the screen.

11. From Visual Studio, open the project and Constants.cs file. In the **AppModeConstants** class, find and update the string values of **ClientId** and **ClientSecret** with the Client ID and key values from Azure management portal. Update your **TenantName** for the authString value (e.g. contoso.onMicrosoft.com). Finally, update the **TenantId** value. Your tenant ID can be discovered by opening the following metadata.xml document: https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml  - replace "graphDir1.onMicrosoft.com", with your tenant's domain value (any domain that is owned by the tenant will work).  The tenantId is a guid, that is part of the sts URL, returned in the first xml node's sts url ("EntityDescriptor"): e.g. "https://sts.windows.net/<tenantIdvalue>"

12. Now Configure a 2nd application object to run the update portion of this app. Repeat steps 1-3, but select **Native Client Application** this time. Give it a name and use **https://localhost/** as the redirect URI. Once again, go into Required permissions and select **Windows Azure Active Directory**, then select **Access the directory as the signed-in user**. Finally, copy the Application ID value for later. Save your changes.

13. Open the Constants.cs file, and replace the **ClientI** in **UserModeConstants** with the Application ID value from the previous step.

14. Build and run your application - you will need to authenticate with valid tenant administrator credentials for your company when you run the application (required for the Create/Update/delete operations).
