using Microsoft.Extensions.Hosting;
using PnP.Core.Services.Builder.Configuration;
using PnP.Core.Auth.Services.Builder.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System.Security.Cryptography.X509Certificates;
using System.Configuration;
using Microsoft.Graph;
using PnP.Core.Services;
using System;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using PnP.Framework;

namespace PnPCore.NET_Framework_4._8.csproj.PnPHost
{
    public class CreatePnPHost
    {
        /// <summary>
        /// PnP Host
        /// </summary>
        public IHost PnPIHost { get; private set; }

        public CreatePnPHost(string site)
        {
            Dictionary<string, string> configuration = new Dictionary<string, string>();
            foreach (var key in ConfigurationManager.AppSettings)
            {
                configuration.Add(key.ToString(), ConfigurationManager.AppSettings[key.ToString()]);
            }

            PnPIHost = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    services.AddLogging(builder =>
                    {
                        builder.AddFilter("Microsoft", LogLevel.Warning)
                               .AddFilter("System", LogLevel.Warning)
                               .AddFilter("PnP.Core.Auth", LogLevel.Warning)
                               .AddConsole();
                    });

                    services.AddPnPCore(options =>
                    {
                        options.PnPContext.GraphFirst = true;
                        options.HttpRequests.UserAgent = $"ISV|{configuration["tenantUri"]}|Product";
                        options.Sites.Add(site, new PnPCoreSiteOptions
                        {
                            SiteUrl = $"https://{configuration["tenantUri"]}.sharepoint.com/sites/{site}"
                        });
                    });
                    services.AddPnPCoreAuthentication(
                        options =>
                        {
                            options.Credentials.Configurations.Add("x509certificate", new PnPCoreAuthenticationCredentialConfigurationOptions
                            {
                                ClientId = configuration["clientId"],
                                TenantId = configuration["tenantId"],
                                X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                                {
                                    StoreName = StoreName.My,
                                    StoreLocation = StoreLocation.LocalMachine,
                                    Thumbprint = configuration["x509thumbprint"]
                                }
                            });
                            options.Credentials.DefaultConfiguration = "x509certificate";
                            options.Sites.Add(site, new PnPCoreAuthenticationSiteOptions
                            {
                                AuthenticationProviderName = "x509certificate"
                            });
                        }
                    );
                })
                .UseConsoleLifetime()
                .Build();
        }

        public static GraphServiceClient CreateGraphClient(PnPContext context)
        {

            // Create a Graph Service client and perform a Graph call using the Microsoft Graph .NET SDK
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
