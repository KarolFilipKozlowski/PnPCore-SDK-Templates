using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using PnP.Core.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services.Builder.Configuration;
using PnP.Framework;
using System;
using Microsoft.Kiota.Abstractions.Authentication;

namespace PnPCore_.NET_7._0.PnPHost
{
    public class CreatePnPHost
    {
        /// <summary>
        /// PnP Host
        /// </summary>
        public IHost PnPIHost { get; private set; }

        public CreatePnPHost()
        {
            PnPIHost = Host.CreateDefaultBuilder()
            .ConfigureServices((pnpContext, services) =>
            {
                services.AddLogging(builder =>
                {
                    builder.AddFilter("Microsoft", LogLevel.Warning)
                           .AddFilter("System", LogLevel.Warning)
                           .AddFilter("PnP.Core.Auth", LogLevel.Warning)
                           .AddConsole();
                });

                // Add PnP Core SDK
                services.AddPnPCore();
                services.Configure<PnPCoreOptions>(pnpContext.Configuration.GetSection("PnPCore"));

                // Add the PnP Core SDK Authentication Providers
                services.AddPnPCoreAuthentication();
                services.Configure<PnPCoreAuthenticationOptions>(pnpContext.Configuration.GetSection("PnPCore"));
            })
            .UseConsoleLifetime()
            .Build();
        }

        public static GraphServiceClient GetGraphContext(PnPContext context)
        {
            var pnpContextAccessToken = (context.AuthenticationProvider.GetAccessTokenAsync(new Uri("https://graph.microsoft.com"))).Result;
            var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(pnpContextAccessToken));
            return new GraphServiceClient(authenticationProvider);
        }

        public static ClientContext GetCSOMContext(PnPContext context)
        {
            return PnPCoreSdk.Instance.GetClientContext(context);
        }
    }

    public class TokenProvider : IAccessTokenProvider
    {
        public TokenProvider(string token) => Token = token;
        public string Token { get; set; }
        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
            CancellationToken cancellationToken = default)
        {
            // get the token and return it in your own way
            return Task.FromResult(Token);
        }

        public AllowedHostsValidator AllowedHostsValidator { get; }
    }
}
