#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Data.Services.Client;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.WebSockets;
using Microsoft.Data.OData;
using Microsoft.Data.Edm;
using Microsoft.Data.Spatial;

#endregion

namespace GraphConsoleAppV3
{
    internal class Requests
    {
        public static ActiveDirectoryClient client;

        public static async Task UserMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Graph Client for user

            //*********************************************************************
            // setup Microsoft Graph Client for user...
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsUser(); ;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            #endregion

            Console.WriteLine("\nStarting user-mode requests...");
            Console.WriteLine("\n=============================\n\n");

            //*********************************************************************
            // The following section may be run by any user, as long as the app
            // has been granted the minimum of User.Read (and User.ReadWrite to update photo)
            // and User.ReadBasic.All scope permissions. Directory.ReadWrite.All
            // or Directory.AccessAsUser.All will also work, but are much more privileged.
            //*********************************************************************
            #region TenantDetails

            //*********************************************************************
            // Get Tenant Details
            // Note: update the string TenantId with your TenantId.
            // This can be retrieved from the login Federation Metadata end point:             
            // https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml
            //  Replace "GraphDir1.onMicrosoft.com" with any domain owned by your organization
            // The returned value from the first xml node "EntityDescriptor", will have a STS URL
            // containing your TenantId e.g. "https://sts.windows.net/4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34/" is returned for GraphDir1.onMicrosoft.com
            //*********************************************************************
            VerifiedDomain initialDomain = new VerifiedDomain();
            VerifiedDomain defaultDomain = new VerifiedDomain();

            ITenantDetail tenant = null;
            Console.WriteLine("\n Retrieving Tenant Details");

            try
            {
                IPagedCollection<ITenantDetail> tenantsCollection = await client.TenantDetails
                    .Where(tenantDetail => tenantDetail.ObjectId.Equals(Constants.TenantId))
                    .ExecuteAsync();
                List<ITenantDetail> tenantsList = tenantsCollection.CurrentPage.ToList();

                if (tenantsList.Count > 0)
                {
                    tenant = tenantsList.First();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting TenantDetails {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (tenant == null)
            {
                Console.WriteLine("Tenant not found");
            }
            else
            {
                TenantDetail tenantDetail = (TenantDetail)tenant;
                Console.WriteLine("Tenant Display Name: " + tenantDetail.DisplayName);

                // Get the Tenant's Verified Domains 
                initialDomain = tenantDetail.VerifiedDomains.First(x => x.Initial.HasValue && x.Initial.Value);
                Console.WriteLine("Initial Domain Name: " + initialDomain.Name);
                defaultDomain = tenantDetail.VerifiedDomains.First(x => x.@default.HasValue && x.@default.Value);
                Console.WriteLine("Default Domain Name: " + defaultDomain.Name);

                // Get Tenant's Tech Contacts
                foreach (string techContact in tenantDetail.TechnicalNotificationMails)
                {
                    Console.WriteLine("Tenant Tech Contact: " + techContact);
                }
            }

            #endregion

            #region Get signed user info, get their photo, and update their photo

            #region Get signed in user details
            User signedInUser = new User();
            try
            {
                signedInUser = (User) await client.Me.ExecuteAsync();
                Console.WriteLine("\nUser UPN: {0}, DisplayName: {1}", signedInUser.UserPrincipalName, signedInUser.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting signed in user {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }


            #endregion

            #region get signed in user's photo
            if (signedInUser.ObjectId != null)
            {
                IUser sUser = (IUser)signedInUser;
                IStreamFetcher photo = (IStreamFetcher)sUser.ThumbnailPhoto;
                try
                {
                    DataServiceStreamResponse response =
                    await photo.DownloadAsync();
                    Console.WriteLine("\nUser {0} GOT thumbnailphoto", signedInUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting the user's photo - may not exist {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }
            #endregion

            #region update signed in user's photo
            // NOTE:  Updating the signed in user's photo requires User.ReadWrite (when available) or 
            // Directory.ReadWrite.All or Directory.AccessAsUser.All
            if (signedInUser.ObjectId != null)
            {
                Console.WriteLine("\nDo you want to update your thumbnail photo with a default icon? Click y/n\n");
                ConsoleKeyInfo key = Console.ReadKey();

                if (key.KeyChar == 'y')
                {
                    IUser sUser = (IUser)signedInUser;

                    //TODO - update with allowed art and save locally with project
                    FileStream fileStream = new FileStream(@"C:\Users\CALVLI\Pictures\profile.PNG", FileMode.Open, FileAccess.Read);
                    try
                    {
                        await sUser.ThumbnailPhoto.UploadAsync((Stream)fileStream, "application/image");
                        Console.WriteLine("\nUser {0} was updated with a thumbnailphoto", signedInUser.DisplayName);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("\nError Updating the user photo {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                    }
                }
            }

            #endregion

            #endregion

            #region People Picker example
            //*********************************************************************
            // People picker
            // Search for a user using text string "Us" match against userPrincipalName, displayName, giveName, surname
            // Requires minimum of User.ReadBasic.All.
            //*********************************************************************
            Console.WriteLine("\nSearching for any user based on UPN, DisplayName, First or Last Name");
            Console.WriteLine("\nPlease enter the user's name you are looking for:");
            String searchString = Console.ReadLine();

            List<IUser> usersList = null;
            IPagedCollection<IUser> searchResults = null;
            try
            {
                IUserCollection userCollection = client.Users;
                searchResults = await userCollection.Where(user =>
                    user.UserPrincipalName.StartsWith(searchString) ||
                    user.DisplayName.StartsWith(searchString) ||
                    user.GivenName.StartsWith(searchString) ||
                    user.Surname.StartsWith(searchString)).Take(10).ExecuteAsync();
                usersList = searchResults.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting User {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (usersList != null && usersList.Count > 0)
            {
                do
                {
                    int j = 1;
                    usersList = searchResults.CurrentPage.ToList();
                    foreach (IUser user in usersList)
                    {
                        Console.WriteLine("User {2} DisplayName: {0}  UPN: {1}",
                            user.DisplayName, user.UserPrincipalName, j);
                        j++;
                    }
                    searchResults = await searchResults.GetNextPageAsync();
                } while (searchResults != null);
            }
            else
            {
                Console.WriteLine("User not found");
            }

            #endregion

            #region Search for user by UPN
            // search for a single user by UPN - the admin user
            searchString = "admin@" + initialDomain.Name;
            Console.WriteLine("\n Retrieving user with UPN {0}", searchString);
            User retrievedUser = new User();
            List<IUser> retrievedUsers = null;
            try
            {
                IPagedCollection<IUser> userCollection = await client.Users
                    .Where(user => user.UserPrincipalName.Equals(searchString))
                    .ExecuteAsync();
                retrievedUsers = userCollection.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting new user {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            // should only find one user with the specified UPN
            if (retrievedUsers != null && retrievedUsers.Count == 1)
            {
                retrievedUser = (User)retrievedUsers.First();
            }
            else
            {
                Console.WriteLine("User not found {0}", searchString);
            }

            #endregion

            #region Get signed in user's manager and direct reports, group memberships and role memberships
            // ***********************************************************************
            // NOTE:  This requires User.Read.All permission scope, or Directory.Read.All or Directory.AccessAsUser.All
            // Group membership requires Group.Read.All or Directory.Read.All (the latter is required for role memberships)
            // Code snippet also demonstrates paging through user's direct reports
            // ***********************************************************************
            
            // manager and reports...
            try
            {
                Console.WriteLine("\nRetrieving signed in user's Manager and Direct Reports");
                IUserFetcher userFetcher = signedInUser as IUserFetcher;
                IDirectoryObject manager = await userFetcher.Manager.ExecuteAsync();
                IPagedCollection<IDirectoryObject> reports = await userFetcher.DirectReports.ExecuteAsync();

                if (manager is User)
                {
                    Console.WriteLine("\n  Manager (user):" + ((IUser)(manager)).DisplayName);
                }
                else if (manager is Contact)
                {
                    Console.WriteLine("\n  Manager (contact):" + ((IContact)(manager)).DisplayName);
                }
                else
                {
                    Console.WriteLine("\n  Manager:  No manager :)");
                }

                if (reports != null)
                {
                    Console.WriteLine("\n  Direct reports:");
                }
                do
                {
                    List<IDirectoryObject> directoryObjects = reports.CurrentPage.ToList();
                    foreach (IDirectoryObject directoryObject in directoryObjects)
                    {
                        if (directoryObject is User)
                        {
                            Console.WriteLine("\n    " + ((IUser)(manager)).DisplayName);
                        }
                        else if (directoryObject is Contact)
                        {
                            Console.WriteLine("\n    " + ((IContact)(manager)).DisplayName);
                        }
                        
                    }
                    reports = await reports.GetNextPageAsync();
                } while (reports != null);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting user's manager and reports {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            // group and role memberships
            Console.WriteLine("\n Signed in user {0} is a member of the following Group and Roles (IDs)", signedInUser.DisplayName);
            IUserFetcher signedInUserFetcher = signedInUser;
            try
            {
                IPagedCollection<IDirectoryObject> pagedCollection = await signedInUserFetcher.MemberOf.ExecuteAsync();
                do
                {
                    List<IDirectoryObject> directoryObjects = pagedCollection.CurrentPage.ToList();
                    foreach (IDirectoryObject directoryObject in directoryObjects)
                    {
                        if (directoryObject is Group)
                        {
                            Group group = directoryObject as Group;
                            Console.WriteLine(" Group: {0}  Description: {1}", group.DisplayName, group.Description);
                        }
                        if (directoryObject is DirectoryRole)
                        {
                            DirectoryRole role = directoryObject as DirectoryRole;
                            Console.WriteLine(" Role: {0}  Description: {1}", role.DisplayName, role.Description);
                        }
                    }
                    pagedCollection = await pagedCollection.GetNextPageAsync();
                } while (pagedCollection != null);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Signed in user's groups and roles memberships. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion

            #region Create a new User
            // **********************************************************
            // Requires Directory.ReadWrite.All or Directory.AccessAsUser.All, and the signed in user
            // must be a privileged user (like a company or user admin)
            // **********************************************************

            IUser newUser = new User();
            if (defaultDomain.Name != null)
            {
                Console.WriteLine("\nCreating a new user...");
                Console.WriteLine("\n  Please enter first name for new user:");
                String firstName = Console.ReadLine();
                Console.WriteLine("\n  Please enter last name for new user:");
                String lastName = Console.ReadLine();
                newUser.DisplayName = firstName + " " + lastName;
                newUser.UserPrincipalName = firstName + "." + lastName + Helper.GetRandomString(4) + "@" + defaultDomain.Name;
                newUser.AccountEnabled = true;
                newUser.MailNickname = firstName + lastName;
                newUser.PasswordProfile = new PasswordProfile
                {
                    Password = "ChangeMe123!",
                    ForceChangePasswordNextLogin = true
                };
                newUser.UsageLocation = "US";
                try
                {
                    await client.Users.AddUserAsync(newUser);
                    Console.WriteLine("\nNew User {0} was created", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError creating new user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Update newly created User properties

            //*******************************************************************************************
            // update the newly created user's Password, PasswordPolicies and City
            //*********************************************************************************************
            if (newUser.ObjectId != null)
            {
                // update User's info
                newUser.City = "Seattle";
                newUser.Country = "UK";
                newUser.Mobile = "+4477889456789";
                newUser.UserType = "Member";

                try
                {
                    await newUser.UpdateAsync();
                    Console.WriteLine("\nUser {0} was updated", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError Updating the user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Reset password for newly created user
            //*******************************************************************************************
            // update the newly created user's Password and PasswordPolicies
            // requires Directory.AccessAsUser.All and that the current user is a user, helpdesk or company admin
            //*********************************************************************************************
            if (newUser.ObjectId != null)
            {
                // update User's password policy and reset password - forcing change password at next logon
                PasswordProfile PasswordProfile = new PasswordProfile
                {
                    Password = "changeMe!",
                    ForceChangePasswordNextLogin = true
                };
                newUser.PasswordProfile = PasswordProfile;
                newUser.PasswordPolicies = "DisablePasswordExpiration, DisableStrongPassword";
                try
                {
                    await newUser.UpdateAsync();
                    Console.WriteLine("\nUser {0} password and policy was reset", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError Updating the user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Other user write operations on the newly created user
            // *************************************************************
            // These operations require more privileged permissions like Directory.ReadWrite.All or Directory.AccessAsUser.All
            // Update signed in user's manager, update group membership
            // **************************************************************

            #region Assign a manager
            // Assign the newly created user a new manager (the signed in user).
            if (newUser.ObjectId != null)
            {
                Console.WriteLine("\n Assign User {0}, {1} as Manager.", signedInUser.DisplayName,
                        newUser.DisplayName);
                newUser.Manager = signedInUser as DirectoryObject;
                try
                {
                    await newUser.UpdateAsync();
                    Console.Write("User {1} is successfully assigned {0} as their Manager.", signedInUser.DisplayName,
                        newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError assigning manager to user. {0} {1}", e.Message,
                            e.InnerException != null ? e.InnerException.Message : "");
                }
            }
            else
                Console.WriteLine("\n Assigning manager failed, because new user was not created.");
            #endregion

            #region Add the new user to a selected group
            // Search for a group and assign the newUser to the found group
           
            //*********************************************************************
            // Search for a group using a startsWith filter (displayName property)
            //*********************************************************************
            Group retrievedGroup = new Group();
            Console.WriteLine("\nSearch for a group to add the current user to:");
            searchString = Console.ReadLine();
            int pickedGroupIndex = -1;

            List<IGroup> foundGroups = null;
            try
            {
                IPagedCollection<IGroup> groupsCollection = await client.Groups
                    .Where(group => group.DisplayName.StartsWith(searchString))
                    .ExecuteAsync();
                foundGroups = groupsCollection.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Group {0} {1}",
                    e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (foundGroups != null && foundGroups.Count > 0)
            {   
                for(int i = 0; i < foundGroups.Count; i++) {
                    Console.WriteLine("\n{0}. {1}", i+1, foundGroups[i].DisplayName);
                }
                Console.WriteLine("Pick the group you want to add the new user to, but entering a number:");
                ConsoleKeyInfo key = Console.ReadKey();

                if (((int) key.KeyChar - 1) < foundGroups.Count) {
                    pickedGroupIndex = (int) key.KeyChar - 1;
                }
            }
            else
            {
                Console.WriteLine("Group Not Found based on search criteria, and hence user not added to group");
            }

            // add new user to picked group
            if (pickedGroupIndex > -1 && foundGroups[pickedGroupIndex].ObjectId != null)
            {
                try
                {
                    retrievedGroup = (Group) foundGroups[pickedGroupIndex];
                    retrievedGroup.Members.Add(newUser as DirectoryObject);
                    await retrievedGroup.UpdateAsync();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError assigning member to group. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }

            }
            #endregion       

            #region Get Group members

            if (retrievedGroup.ObjectId != null)
            {
                Console.WriteLine("\n Enumerating group members for: " + retrievedGroup.DisplayName + "\n " + retrievedGroup.Description);

                //*********************************************************************
                // get the groups' membership - 
                // Note this method retrieves ALL links in one request - please use this method with care - this
                // may return a very large number of objects
                //*********************************************************************
                IGroupFetcher retrievedGroupFetcher = retrievedGroup;
                try
                {
                    IPagedCollection<IDirectoryObject> members = await retrievedGroupFetcher.Members.ExecuteAsync();
                    Console.WriteLine(" Members:");
                    do
                    {
                        List<IDirectoryObject> directoryObjects = members.CurrentPage.ToList();
                        foreach (IDirectoryObject member in directoryObjects)
                        {
                            if (member is User)
                            {
                                User user = member as User;
                                Console.WriteLine("User DisplayName: {0}  UPN: {1}",
                                    user.DisplayName,
                                    user.UserPrincipalName);
                            }
                            if (member is Group)
                            {
                                Group group = member as Group;
                                Console.WriteLine("Group DisplayName: {0}", group.DisplayName);
                            }
                            if (member is Contact)
                            {
                                Contact contact = member as Contact;
                                Console.WriteLine("Contact DisplayName: {0}", contact.DisplayName);
                            }
                        }
                        members = await members.GetNextPageAsync();
                    } while (members != null);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting groups' membership. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Remove new user from the  group

            //*********************************************************************************************
            // Delete user from the earlier selected Group 
            //*********************************************************************************************
            if (retrievedGroup.ObjectId != null)
            {
                try
                {
                    retrievedGroup.Members.Remove(newUser as DirectoryObject);
                    await retrievedGroup.UpdateAsync();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nRemoving user from group failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region New user License Assignment

            //*********************************************************************************************
            // User License Assignment - assign EnterprisePack license to new user, and disable SharePoint service
            //   first get a list of Tenant's subscriptions and find the "Enterprisepack" one
            //   Enterprise Pack includes service Plans for ExchangeOnline, SharePointOnline and LyncOnline
            //   validate that Subscription is Enabled and there are enough units left to assign to users
            //*********************************************************************************************
            IPagedCollection<ISubscribedSku> skus = null;
            try
            {
                skus = await client.SubscribedSkus.ExecuteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Applications {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            if (skus != null)
            {
                do
                {
                    List<ISubscribedSku> subscribedSkus = skus.CurrentPage.ToList();
                    foreach (ISubscribedSku sku in subscribedSkus)
                    {
                        if (sku.SkuPartNumber == "ENTERPRISEPACK")
                        {
                            if ((sku.PrepaidUnits.Enabled.Value > sku.ConsumedUnits) &&
                                (sku.CapabilityStatus == "Enabled"))
                            {
                                // create addLicense object and assign the Enterprise Sku GUID to the skuId
                                AssignedLicense addLicense = new AssignedLicense { SkuId = sku.SkuId.Value };

                                // find plan id of SharePoint Service Plan
                                foreach (ServicePlanInfo servicePlan in sku.ServicePlans)
                                {
                                    if (servicePlan.ServicePlanName.Contains("SHAREPOINT"))
                                    {
                                        addLicense.DisabledPlans.Add(servicePlan.ServicePlanId.Value);
                                        break;
                                    }
                                }

                                IList<AssignedLicense> licensesToAdd = new[] { addLicense };
                                IList<Guid> licensesToRemove = new Guid[] { };

                                // attempt to assign the license object to the new user 
                                try
                                {
                                    if (newUser.ObjectId != null)
                                    {
                                        await newUser.AssignLicenseAsync(licensesToAdd, licensesToRemove);
                                        Console.WriteLine("\n User {0} was assigned license {1}",
                                            newUser.DisplayName,
                                            addLicense.SkuId);
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("\nLicense assingment failed {0} {1}", e.Message,
                                        e.InnerException != null ? e.InnerException.Message : "");
                                }
                            }
                        }
                    }
                    skus = await skus.GetNextPageAsync();
                } while (skus != null);
            }

            #endregion

            #endregion

            #region Create a new Group

            //*********************************************************************************************
            // Create a new Group
            //*********************************************************************************************
            Group newGroup = new Group
            {
                DisplayName = "newGroup" + Helper.GetRandomString(8),
                Description = "Best Group ever",
                MailNickname = "group" + Helper.GetRandomString(4),
                MailEnabled = false,
                SecurityEnabled = true
            };
            try
            {
                await client.Groups.AddGroupAsync(newGroup);
                Console.WriteLine("\nNew Group {0} was created", newGroup.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError creating new Group {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get All Roles

            //*********************************************************************
            // Get All Roles
            //*********************************************************************
            List<IDirectoryRole> foundRoles = null;
            try
            {
                IPagedCollection<IDirectoryRole> rolesCollection = await client.DirectoryRoles.ExecuteAsync();
                foundRoles = rolesCollection.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Roles {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (foundRoles != null && foundRoles.Count > 0)
            {
                foreach (IDirectoryRole role in foundRoles)
                {
                    Console.WriteLine("\n Found Role: {0} {1} {2} ",
                        role.DisplayName, role.Description, role.ObjectId);
                }
            }
            else
            {
                Console.WriteLine("Role Not Found {0}", searchString);
            }

            #endregion

            #region Get Service Principals
            string graphAppObjectId = "";

            //*********************************************************************
            // get the Service Principals
            //*********************************************************************
            IPagedCollection<IServicePrincipal> servicePrincipals = null;
            try
            {
                servicePrincipals = await client.ServicePrincipals.ExecuteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Service Principal {0} {1}",
                    e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (servicePrincipals != null)
            {
                do
                {
                    List<IServicePrincipal> servicePrincipalsList = servicePrincipals.CurrentPage.ToList();
                    foreach (IServicePrincipal servicePrincipal in servicePrincipalsList)
                    {
                        Console.WriteLine("Service Principal AppId: {0}  Name: {1}", servicePrincipal.AppId,
                            servicePrincipal.DisplayName);
                        // find the Graph API service principal objectId
                        if (servicePrincipal.AppId == "00000002-0000-0000-c000-000000000000")
                        {
                            graphAppObjectId = servicePrincipal.ObjectId;
                        }
                    }
                    servicePrincipals = await servicePrincipals.GetNextPageAsync();
                } while (servicePrincipals != null);
            }

            #endregion

            #region Get Applications

            //*********************************************************************
            // get the Application objects
            //*********************************************************************
            IPagedCollection<IApplication> applications = null;
            try
            {
                applications = await client.Applications.Take(50).ExecuteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Applications {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            if (applications != null)
            {
                do
                {
                    List<IApplication> appsList = applications.CurrentPage.ToList();
                    foreach (IApplication app in appsList)
                    {
                        Console.WriteLine("Application AppId: {0}  Name: {1}", app.AppId, app.DisplayName);
                    }
                    applications = await applications.GetNextPageAsync();
                } while (applications != null);
            }

            #endregion

            #region Create Application

            //*********************************************************************************************
            // Create a new Application object, with an App Role definition
            //*********************************************************************************************
            Application newApp = new Application { DisplayName = "Test-Demo App " + Helper.GetRandomString(4) };
            newApp.IdentifierUris.Add("https://localhost/demo/" + Guid.NewGuid());
            newApp.ReplyUrls.Add("https://localhost/demo");
            AppRole appRole = new AppRole()
            {
                Id = Guid.NewGuid(),
                IsEnabled = true,
                DisplayName = "Something",
                Description = "Anything",
                Value = "policy.write"
            };
            appRole.AllowedMemberTypes.Add("User");
            newApp.AppRoles.Add(appRole);

            // Add a password key
            PasswordCredential password = new PasswordCredential
            {
                StartDate = DateTime.UtcNow,
                EndDate = DateTime.UtcNow.AddYears(1),
                Value = "password",
                KeyId = Guid.NewGuid()
            };
            newApp.PasswordCredentials.Add(password);

            try
            {
                await client.Applications.AddApplicationAsync(newApp);
                Console.WriteLine("New Application created: " + newApp.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("Application Creation execption: {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            // add an owner for the newly created application
            newApp.Owners.Add(newUser as DirectoryObject);
            try
            {
                await newApp.UpdateAsync();
                Console.WriteLine("Added owner: " + newApp.DisplayName, newUser.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("Application owner addition execption: {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            // check the ownership for the newly created application
            try
            {
                IApplication appCheck = await client.Applications.GetByObjectId(newApp.ObjectId).ExecuteAsync();
                IApplicationFetcher appCheckFetcher = appCheck as IApplicationFetcher;

                IPagedCollection<IDirectoryObject> appOwners = await appCheckFetcher.Owners.ExecuteAsync();

                do
                {
                    List<IDirectoryObject> directoryObjects = appOwners.CurrentPage.ToList();
                    foreach (IDirectoryObject directoryObject in directoryObjects)
                    {
                        if (directoryObject is User)
                        {
                            User appOwner = directoryObject as User;
                            Console.WriteLine("Application {0} has {1} as owner", appCheck.DisplayName, appOwner.DisplayName);
                        }
                    }
                    appOwners = await appOwners.GetNextPageAsync();
                } while (appOwners != null);
            }
            catch (Exception e)
            {
                Console.WriteLine("Application owner check exception: {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion

            #region Create Service Principal

            //*********************************************************************************************
            // create a new Service principal, from the application object that was just created
            //*********************************************************************************************
            ServicePrincipal newServicePrincpal = new ServicePrincipal();
            if (newApp.AppId != null)
            {
                newServicePrincpal.DisplayName = newApp.DisplayName;
                newServicePrincpal.AccountEnabled = true;
                newServicePrincpal.AppId = newApp.AppId;
                try
                {
                    await client.ServicePrincipals.AddServicePrincipalAsync(newServicePrincpal);
                    Console.WriteLine("New Service Principal created: " + newServicePrincpal.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Service Principal Creation execption: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Working with directory schema extensions
            string extName = "linkedInUserId";

            #region Create an Extension Property
            // **************************************************************************************************
            // Create a new extension property - to extend the user entity
            // This is accomplished by declaring the extension property through an application object
            // **************************************************************************************************
            if (newApp.ObjectId != null)
            {
                ExtensionProperty linkedInUserId = new ExtensionProperty
                {
                    Name = extName,
                    DataType = "String",
                    TargetObjects = { "User" }
                };
                try
                {
                    // firstly, let's write out all the existing cloud extension properties in the tenant
                    IEnumerable<IExtensionProperty> allExts = await client.GetAvailableExtensionPropertiesAsync(false);
                    foreach (ExtensionProperty ext in allExts)
                    {
                        Console.WriteLine("\nExtension: {0}", ext.Name);
                    }
                    newApp.ExtensionProperties.Add(linkedInUserId);
                    await newApp.UpdateAsync();
                    Console.WriteLine("\nUser object extended successfully with extension: {0}.", linkedInUserId.Name);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError extending the user object {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            #endregion

            #region Manipulate an Extension Property
                // **************************************************************************************************
                // Update an extension property that exists on the user entity
                // **************************************************************************************************
            
                // TODO TODO - need to fix manipulating and getting extension value.

                // create the extension attribute name, for the extension that we just created
                string attributeName = "extension_" + newApp.AppId + "_" + extName;
                try
                {
                    if (retrievedUser != null && retrievedUser.ObjectId != null)
                    {
                        retrievedUser.SetExtendedProperty(attributeName, "user@linkedin.com");
                        await retrievedUser.UpdateAsync();
                        Console.WriteLine("\nUser {0}'s extended property set successully.", retrievedUser.DisplayName);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError Updating the user object {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

                #endregion

            #region Get an Extension Property

                try
                {
                    if (retrievedUser != null && retrievedUser.ObjectId != null)
                    {
                        IReadOnlyDictionary<string, object> extendedProperties = retrievedUser.GetExtendedProperties();
                        object extendedProperty = extendedProperties[attributeName];
                        Console.WriteLine("\n Retrieved User {0}'s extended property value is: {1}.", retrievedUser.DisplayName,
                            extendedProperty);
                    }       
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError Updating the user object {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

                #endregion

            #endregion

            #region Assign an app role
                try
                {
                    User user =
                            (User) (await client.Users.ExecuteAsync()).CurrentPage.ToList().FirstOrDefault();
                    if (newApp.ObjectId != null && user != null && newServicePrincpal.ObjectId != null)
                    {
                        // create the app role assignment
                        AppRoleAssignment appRoleAssignment = new AppRoleAssignment();
                        appRoleAssignment.Id = appRole.Id;
                        appRoleAssignment.ResourceId = Guid.Parse(newServicePrincpal.ObjectId);
                        appRoleAssignment.PrincipalType = "User";
                        appRoleAssignment.PrincipalId = Guid.Parse(user.ObjectId);
                        user.AppRoleAssignments.Add(appRoleAssignment);
                        
                        // assign the app role
                        await user.UpdateAsync();
                        Console.WriteLine("User {0} is successfully assigned an app (role).", retrievedUser.DisplayName);

                        // remove the app role
                        user.AppRoleAssignments.Remove(appRoleAssignment);
                        await user.UpdateAsync();
                        Console.WriteLine("User {0} is successfully UNassigned an app (role).", retrievedUser.DisplayName);

                    }
                }

                catch (Exception e)
                {
                    Console.WriteLine("Direct Permission Assignment failed: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

            }
                #endregion

            #region Get Devices

            //*********************************************************************************************
            // Get a list of Mobile Devices from tenant
            //*********************************************************************************************
            Console.WriteLine("\nGetting Devices");
            IPagedCollection<IDevice> devices = null;
            try
            {
                devices = await client.Devices.ExecuteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine("/nError getting devices {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (devices != null)
            {
                do
                {
                    List<IDevice> devicesList = devices.CurrentPage.ToList();
                    foreach (IDevice device in devicesList)
                    {
                        if (device.ObjectId != null)
                        {
                            Console.WriteLine("Device ID: {0}, Type: {1}", device.DeviceId, device.DeviceOSType);
                            IPagedCollection<IDirectoryObject> registeredOwners = device.RegisteredOwners;
                            if (registeredOwners != null)
                            {
                                do
                                {
                                    List<IDirectoryObject> registeredOwnersList = registeredOwners.CurrentPage.ToList();
                                    foreach (IDirectoryObject owner in registeredOwnersList)
                                    {
                                        Console.WriteLine("Device Owner ID: " + owner.ObjectId);
                                    }
                                    registeredOwners = await registeredOwners.GetNextPageAsync();
                                } while (registeredOwners != null);
                            }
                        }
                    }
                    devices = await devices.GetNextPageAsync();
                } while (devices != null);
            }

            #endregion

            #region Create a new consentable OAuth2 permission

            //*********************************************************************************************
            // Create new  oauth2 permission object
            //*********************************************************************************************
            if (newServicePrincpal.ObjectId != null)
            {
                OAuth2PermissionGrant permissionObject = new OAuth2PermissionGrant();
                permissionObject.ConsentType = "AllPrincipals";
                permissionObject.Scope = "user_impersonation";
                permissionObject.StartTime = DateTime.Now;
                permissionObject.ExpiryTime = (DateTime.Now).AddMonths(12);

                // resourceId is objectId of the resource, in this case objectId of AzureAd (Graph API)
                permissionObject.ResourceId = graphAppObjectId;

                //ClientId = objectId of servicePrincipal
                permissionObject.ClientId = newServicePrincpal.ObjectId;

                // add the oauth2 permission scope grant
                try
                {
                    await client.Oauth2PermissionGrants.AddOAuth2PermissionGrantAsync(permissionObject);
                    Console.WriteLine("New Permission object created: " + permissionObject.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Permission Creation exception: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }

                // remove the oauth2 permission scope grant
                try
                {
                    newServicePrincpal.Oauth2PermissionGrants.Remove(permissionObject);
                    await newServicePrincpal.UpdateAsync();
                    Console.WriteLine("Removed Permission object: " + permissionObject.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Permission Creation exception: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }

                try
                {
                    await newServicePrincpal.DeleteAsync();
                    Console.WriteLine("Deleted service principal object: " + newServicePrincpal.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Service principal Deletion execption: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

            }
            #endregion

            #region Get All Permissions

            //*********************************************************************************************
            // get all Permission Objects
            //*********************************************************************************************
            Console.WriteLine("\n Getting Permissions");
            IPagedCollection<IOAuth2PermissionGrant> permissions = null;
            try
            {
                permissions = await client.Oauth2PermissionGrants.ExecuteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (permissions != null)
            {
                do
                {
                    List<IOAuth2PermissionGrant> perms = permissions.CurrentPage.ToList();
                    foreach (IOAuth2PermissionGrant perm in perms)
                    {
                        Console.WriteLine("Permission: {0}  Name: {1}", perm.ClientId, perm.Scope);
                    }
                    permissions = await permissions.GetNextPageAsync();
                } while (permissions != null);
            }

            #endregion

            #region CleanUp
            #region Delete the new user
            //*********************************************************************************************
            // Delete the user that we just created earlier
            //*********************************************************************************************
            if (newUser.ObjectId != null)
            {
                try
                {
                    await newUser.DeleteAsync();
                    Console.WriteLine("\nUser {0} was deleted", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Deleting User failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Delete new Group

            //*********************************************************************************************
            // Delete the Group that we just created
            //*********************************************************************************************
            if (newGroup.ObjectId != null)
            {
                try
                {
                    await newGroup.DeleteAsync();
                    Console.WriteLine("\nGroup {0} was deleted", newGroup.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Deleting Group failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Delete Application

            //*********************************************************************************************
            // Delete Application Objects
            //*********************************************************************************************
            if (newApp.ObjectId != null)
            {
                try
                {
                    await newApp.DeleteAsync();
                    Console.WriteLine("Deleted Application object: " + newApp.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Application Deletion execption: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion
            #endregion

            #region Batch Operations

            //*********************************************************************************************
            // Show Batching with 3 operators.  Note: up to 5 operations can be in a batch
            //*********************************************************************************************
            IReadOnlyQueryableSet<User> userQuery = client.DirectoryObjects.OfType<User>();
            IReadOnlyQueryableSet<Group> groupsQuery = client.DirectoryObjects.OfType<Group>();
            IReadOnlyQueryableSet<DirectoryRole> rolesQuery =
                client.DirectoryObjects.OfType<DirectoryRole>();
            try
            {
                IBatchElementResult[] batchResult = await 
                    client.Context.ExecuteBatchAsync(userQuery, groupsQuery, rolesQuery);
                int responseCount = 1;
                foreach (IBatchElementResult result in batchResult)
                {
                    if (result.FailureResult != null)
                    {
                        Console.WriteLine("Failed: {0} ",
                            result.FailureResult.InnerException);
                    }
                    if (result.SuccessResult != null)
                    {
                        Console.WriteLine("Batch Item Result {0} succeeded",
                            responseCount++);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Batch execution failed. : {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion


        }
        public static async Task AppMode()
        {
            // record start DateTime of execution
            //string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Microsoft Graph Client for app
            //*********************************************************************
            // setup Microsoft Graph Client for app
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }
            #endregion

            Console.WriteLine("\nStarting app-mode requests...");
            Console.WriteLine("\nAll requests are done in the context of the application only (daemon style app)\n\n");
            Console.WriteLine("\n=============================\n\n");

            #region People Picker
            //*********************************************************************
            // People picker
            // Search for a user using text string "Us" match against userPrincipalName, displayName, giveName, surname
            // Requires minimum of Directory.Read.All (and soon User.Read.All).
            //*********************************************************************
            try
            {
                Console.WriteLine("\nSearching for any user based on UPN, DisplayName, First or Last Name");
                Console.WriteLine("\nPlease enter the user's name you are looking for:");
                String searchString = Console.ReadLine();

                List<IUser> usersList = null;
                IPagedCollection<IUser> searchResults = null;
                try
                {
                    IUserCollection userCollection = client.Users;
                    searchResults = await userCollection.Where(user =>
                        user.UserPrincipalName.StartsWith(searchString) ||
                        user.DisplayName.StartsWith(searchString) ||
                        user.GivenName.StartsWith(searchString) ||
                        user.Surname.StartsWith(searchString)).Take(10).ExecuteAsync();
                    usersList = searchResults.CurrentPage.ToList();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting User {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

                if (usersList != null && usersList.Count > 0)
                {
                    do
                    {
                        int j = 1;
                        usersList = searchResults.CurrentPage.ToList();
                        foreach (IUser user in usersList)
                        {
                            Console.WriteLine("User {2} DisplayName: {0}  UPN: {1}",
                                user.DisplayName, user.UserPrincipalName, j);
                            j++;
                        }
                        searchResults = await searchResults.GetNextPageAsync();
                    } while (searchResults != null);
                }
                else
                {
                    Console.WriteLine("User not found");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion
        }
    }
}