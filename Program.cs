using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Security;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
  public class Program
  {
    public static void Main(string[] args)
    {
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file.");
        return;
      }
      var userName = ReadUsername();
      var userPassword = ReadPassword();

      var client = GetAuthenticatedGraphClient(config, userName, userPassword);
      // request 1 - all groups
      Console.WriteLine("\n\nREQUEST 1 - ALL GROUPS:");
      var requestAllGroups = client.Groups.Request();
      var resultsAllGroups = requestAllGroups.GetAsync().Result;
      foreach (var group in resultsAllGroups)
      {
        Console.WriteLine(group.Id + ": " + group.DisplayName + " <" + group.Mail + ">");
      }

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestAllGroups.GetHttpRequestMessage().RequestUri);

      var groupId = "669c3f6d-2e41-40d1-8fa6-3688e47b3e0c";
      // request 2 - one group
      Console.WriteLine("\n\nREQUEST 2 - ONE GROUP:");
      var requestGroup = client.Groups[groupId].Request();
      var resultsGroup = requestGroup.GetAsync().Result;
      Console.WriteLine(resultsGroup.Id + ": " + resultsGroup.DisplayName + " <" + resultsGroup.Mail + ">");

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestGroup.GetHttpRequestMessage().RequestUri);

      // request 3 - group owners
      Console.WriteLine("\n\nREQUEST 3 - GROUP OWNERS:");
      var requestGroupOwners = client.Groups[groupId].Owners.Request();
      var resultsGroupOwners = requestGroupOwners.GetAsync().Result;
      foreach (var owner in resultsGroupOwners)
      {
        var ownerUser = owner as Microsoft.Graph.User;
        if (ownerUser != null)
        {
          Console.WriteLine(ownerUser.Id + ": " + ownerUser.DisplayName + " <" + ownerUser.Mail + ">");
        }
      }

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestGroupOwners.GetHttpRequestMessage().RequestUri);

      // request 4 - group members
      Console.WriteLine("\n\nREQUEST 4 - GROUP MEMBERS:");
      var requestGroupMembers = client.Groups[groupId].Members.Request();
      var resultsGroupMembers = requestGroupMembers.GetAsync().Result;
      foreach (var member in resultsGroupMembers)
      {
        var memberUser = member as Microsoft.Graph.User;
        if (memberUser != null)
        {
          Console.WriteLine(memberUser.Id + ": " + memberUser.DisplayName + " <" + memberUser.Mail + ">");
        }
      }
      var client1 = GetAuthenticatedGraphClient(config, userName, userPassword);
      var requestOwnerOf = client1.Me.OwnedObjects.Request();
      var resultsOwnerOf = requestOwnerOf.GetAsync().Result;
      foreach (var ownedObject in resultsOwnerOf)
      {
        var group = ownedObject as Microsoft.Graph.Group;
        var role = ownedObject as Microsoft.Graph.DirectoryRole;
        if (group != null)
        {
          Console.WriteLine("Office 365 Group: " + group.Id + ": " + group.DisplayName);
        }
        else if (role != null)
        {
          Console.WriteLine("  Security Group: " + role.Id + ": " + role.DisplayName);
        }
        else
        {
          Console.WriteLine(ownedObject.ODataType + ": " + ownedObject.Id);
        }
      }

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestGroupMembers.GetHttpRequestMessage().RequestUri);
    }
    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
                          .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                          .AddJsonFile("appsettings.json", false, true)
                          .Build();

        if (string.IsNullOrEmpty(config["applicationId"]) ||
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }
    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("User.Read");
      scopes.Add("User.ReadBasic.All");
      scopes.Add("Group.Read.All");
      scopes.Add("Group.ReadWrite.All");
      scopes.Add("Directory.Read.All");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
    }
    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
    private static SecureString ReadPassword()
    {
      Console.WriteLine("Enter your password");
      SecureString password = new SecureString();
      while (true)
      {
        ConsoleKeyInfo c = Console.ReadKey(true);
        if (c.Key == ConsoleKey.Enter)
        {
          break;
        }
        password.AppendChar(c.KeyChar);
        Console.Write("*");
      }
      Console.WriteLine();
      return password;
    }
    private static string ReadUsername()
    {
      string? username;
      Console.WriteLine("Enter your username");
      username = Console.ReadLine();
      return username ?? "";
    }
    private static async Task<Microsoft.Graph.Group> CreateGroupAsync(GraphServiceClient client)
    {
      // create object to define members & owners as 'additionalData'
      var additionalData = new Dictionary<string, object>();
      additionalData.Add("owners@odata.bind",
        new string[] {
      "https://graph.microsoft.com/v1.0/users/d280a087-e05b-4c23-b073-738cdb82b25e"
        }
      );
      additionalData.Add("members@odata.bind",
        new string[] {
      "https://graph.microsoft.com/v1.0/users/70c095fe-df9d-4250-867d-f298e237d681",
      "https://graph.microsoft.com/v1.0/users/8c2da469-1eba-47a4-9322-ee0ddd24d99a"
        }
      );

      var group = new Microsoft.Graph.Group
      {
        AdditionalData = additionalData,
        Description = "My first group created with the Microsoft Graph .NET SDK",
        DisplayName = "My First Group",
        GroupTypes = new List<String>() { "Unified" },
        MailEnabled = true,
        MailNickname = "myfirstgroup01",
        SecurityEnabled = false
      };

      var requestNewGroup = client.Groups.Request();
      return await requestNewGroup.AddAsync(group);
    }
    private static async Task<Microsoft.Graph.Team> TeamifyGroupAsync(GraphServiceClient client, string groupId)
    {
      var team = new Microsoft.Graph.Team
      {
        MemberSettings = new TeamMemberSettings
        {
          AllowCreateUpdateChannels = true,
          ODataType = null
        },
        MessagingSettings = new TeamMessagingSettings
        {
          AllowUserEditMessages = true,
          AllowUserDeleteMessages = true,
          ODataType = null
        },
        ODataType = null
      };

      var requestTeamifiedGroup = client.Groups[groupId].Team.Request();
      return await requestTeamifiedGroup.PutAsync(team);
    }
  }
}