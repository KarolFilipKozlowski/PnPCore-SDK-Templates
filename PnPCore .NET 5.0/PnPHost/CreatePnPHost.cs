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

namespace PnPCore_.NET_5._0.PnPHost
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
            return new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                return context.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
            }));
        }

        public static ClientContext GetCSOMContext(PnPContext context)
        {
            return PnPCoreSdk.Instance.GetClientContext(context);
        }
    }
}
