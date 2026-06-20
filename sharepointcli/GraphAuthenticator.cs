using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;
using System.Diagnostics;

namespace sharepointcli;

internal sealed class GraphAuthenticator
{
    private const string ClientId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
    private const string TenantId = "organizations";

    private readonly IPublicClientApplication _app;

    public GraphAuthenticator()
    {
        BrokerOptions brokerOptions = new BrokerOptions(BrokerOptions.OperatingSystems.Windows);

        _app = PublicClientApplicationBuilder.Create(ClientId)
               .WithTenantId(TenantId)
               .WithBroker(brokerOptions)
               .WithLogging((lvl, msg, _) =>
               {
                   if (lvl == LogLevel.Info)
                   {
                       SystemEventLogger.Info(msg);
                   }
                   else
                   {
                       SystemEventLogger.Error(msg);
                   }
               }, enablePiiLogging: true)
               .WithParentActivityOrWindow(GetTerminalWindow)
               .WithRedirectUri("ms-appx-web://Microsoft.AAD.BrokerPlugin/1fec8e78-bce4-4aaf-ab1b-5451cc387264")
               .Build();

        MsalCacheHelper cacheHelper = CreateCacheHelperAsync().GetAwaiter().GetResult();

        cacheHelper.RegisterCache(_app.UserTokenCache);
    }

    public IntPtr GetTerminalWindow()
    {
        return Process.GetCurrentProcess().Handle;
    }

    public async Task<string> AcquireAccessTokenAsync(string[] scopes)
    {
        AuthenticationResult authResult = null;

        IAccount firstAccount = (await _app.GetAccountsAsync()).FirstOrDefault();

        if (firstAccount == null)
        {
            firstAccount = PublicClientApplication.OperatingSystemAccount;
        }

        try
        {
            authResult = await _app.AcquireTokenSilent(scopes, firstAccount)
                .ExecuteAsync();
        }
        catch (Exception ex)
        {
            SystemEventLogger.Info($"Silent token acquisition failed: {ex.GetType().Name}");
            try
            {
                authResult = await _app.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .ExecuteAsync();
            }
            catch (MsalException msalex)
            {
                SystemEventLogger.Error($"Interactive token acquisition failed: {msalex.Message}");
                Console.WriteLine($"Error Acquiring Token:{System.Environment.NewLine}{msalex}");
            }
        }

        return authResult.AccessToken;
    }

    private static async Task<MsalCacheHelper> CreateCacheHelperAsync()
    {
        //DPAPI
        var storageProperties = new StorageCreationPropertiesBuilder(
                          System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".msalcache.bin", MsalCacheHelper.UserRootDirectory).Build();

        MsalCacheHelper cacheHelper = await MsalCacheHelper.CreateAsync(
                    storageProperties,
                    new TraceSource("MSAL.CacheTrace"))
                 .ConfigureAwait(false);

        return cacheHelper;
    }
}
