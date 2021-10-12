using System;
using System.Linq;
using System.Configuration;
using System.IO;
using System.Security;
using System.Reflection;
using Microsoft.Rest;
using Microsoft.Identity.Client;
using Microsoft.PowerBI.Api;

namespace ExportReportToFile.Models {

  class TokenManager {

    public const string urlPowerBiServiceApiRoot = "https://api.powerbi.com/";
    private const string tenantCommonAuthority = "https://login.microsoftonline.com/organizations";

    private static string applicationId = ConfigurationManager.AppSettings["public-application-id"];
    private static string redirectUri = ConfigurationManager.AppSettings["redirect-uri"];

    private static string confidentialApplicationId = ConfigurationManager.AppSettings["confidential-application-id"];
    private static string confidentialApplicationSecret = ConfigurationManager.AppSettings["confidential-application-secret"];
    private static string tenantName = ConfigurationManager.AppSettings["tenant-name"];
    private readonly static string tenantSpecificAuthority = "https://login.microsoftonline.com/" + tenantName;

    public static string GetAccessTokenInteractive(string[] scopes) {

      // create new public client application
      var appPublic = PublicClientApplicationBuilder.Create(applicationId)
                    .WithAuthority(tenantCommonAuthority)
                    .WithRedirectUri(redirectUri)
                    .Build();

      AuthenticationResult authResult = appPublic.AcquireTokenInteractive(scopes).ExecuteAsync().Result;

      // return access token to caller
      return authResult.AccessToken;
    }

    public static PowerBIClient GetPowerBiClientInteractive(string[] scopes) {

      string accessToken = TokenManager.GetAccessTokenInteractive(scopes);
      TokenCredentials tokenCredentials = new TokenCredentials(accessToken, "Bearer");
      PowerBIClient pbiClient = new PowerBIClient(new Uri(urlPowerBiServiceApiRoot), tokenCredentials);
      return pbiClient;
    }


    public static string GetAccessToken(string[] scopes) {

      // create new public client application
      var appPublic = PublicClientApplicationBuilder.Create(applicationId)
                      .WithAuthority(tenantCommonAuthority)
                      .WithRedirectUri(redirectUri)
                      .Build();

      // connect application to token cache
      TokenCacheHelper.EnableSerialization(appPublic.UserTokenCache);

      AuthenticationResult authResult;
      try {
        // try to acquire token from token cache
        var user = appPublic.GetAccountsAsync().Result.FirstOrDefault();
        authResult = appPublic.AcquireTokenSilent(scopes, user).ExecuteAsync().Result;
      }
      catch {
        authResult = appPublic.AcquireTokenInteractive(scopes).ExecuteAsync().Result;
      }

      // return access token to caller
      return authResult.AccessToken;
    }

    public static PowerBIClient GetPowerBiClient(string[] scopes) {
      var tokenCredentials = new TokenCredentials(GetAccessToken(scopes), "Bearer");
      return new PowerBIClient(new Uri(urlPowerBiServiceApiRoot), tokenCredentials);
    }

    static class TokenCacheHelper {

      private static readonly string CacheFilePath = Assembly.GetExecutingAssembly().Location + ".tokencache.json";
      private static readonly object FileLock = new object();

      public static void EnableSerialization(ITokenCache tokenCache) {
        tokenCache.SetBeforeAccess(BeforeAccessNotification);
        tokenCache.SetAfterAccess(AfterAccessNotification);
      }

      private static void BeforeAccessNotification(TokenCacheNotificationArgs args) {
        lock (FileLock) {
          // repopulate token cache from persisted store
          args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath) ? File.ReadAllBytes(CacheFilePath) : null);
        }
      }

      private static void AfterAccessNotification(TokenCacheNotificationArgs args) {
        // if the access operation resulted in a cache update
        if (args.HasStateChanged) {
          lock (FileLock) {
            // write token cache changes to persistent store
            File.WriteAllBytes(CacheFilePath, args.TokenCache.SerializeMsalV3());
          }
        }
      }
    }

    private static string cachedAppOnlyToken = string.Empty;

    static string GetAppOnlyAccessToken() {

      if (cachedAppOnlyToken.Equals(string.Empty)) {
        var appConfidential = ConfidentialClientApplicationBuilder.Create(confidentialApplicationId)
                                .WithClientSecret(confidentialApplicationSecret)
                                .WithAuthority(tenantSpecificAuthority)
                                .Build();

        string[] scopesDefault = new string[] { "https://analysis.windows.net/powerbi/api/.default" };
        var authResult = appConfidential.AcquireTokenForClient(scopesDefault).ExecuteAsync().Result;
        cachedAppOnlyToken = authResult.AccessToken;
      }
      return cachedAppOnlyToken;
    }
    public static PowerBIClient GetPowerBiAppOnlyClient() {
      var tokenCredentials = new TokenCredentials(GetAppOnlyAccessToken(), "Bearer");
      return new PowerBIClient(new Uri(urlPowerBiServiceApiRoot), tokenCredentials);
    }

  }


}
