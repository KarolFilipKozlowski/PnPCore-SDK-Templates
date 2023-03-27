using PnP.Core.Services;
using PnP.Framework;
using PnPCore.NET_Framework_4._8.csproj.PnPHost;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;

namespace PnPCore.NET_Framework_4._8.csproj
{
    internal class Program
    {
        /// As value set target site from AppSecret.config by it SiteName key:
        private const string SITETARGET = "site";

        public static async Task Main(string[] args)
        {
            // Creates and configures the PnP host
            var host = new CreatePnPHost(site: SITETARGET);
            await host.PnPIHost.StartAsync();
            using (var scope = host.PnPIHost.Services.CreateScope())
            {
                // Create the PnPContext
                // Documentation -> https://pnp.github.io/pnpcore/
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
                using (var context = await pnpContextFactory.CreateAsync(name: SITETARGET))
                {
                    // Load the Title property of the site's root web
                    await context.Web.LoadAsync(p => p.Title);
                    Console.WriteLine($"The title of the web is {context.Web.Title}");

                    #region CSOM Context
                    // Create a CSOM Context
                    // Documentation -> https://pnp.github.io/pnpframework/
                    using (Microsoft.SharePoint.Client.ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(context))
                    {
                        Microsoft.SharePoint.Client.Web web = csomContext.Web;
                        csomContext.Load(web);
                        csomContext.ExecuteQuery();
                        Console.WriteLine($"The title of the web is {web.Title}");
                    }
                    #endregion

                    #region Graph SDK Context
                    // Create a CSOM Context
                    // Documentation -> https://learn.microsoft.com/pl-pl/graph/api/overview?view=graph-rest-1.0
                    Microsoft.Graph.GraphServiceClient graphServiceClient = CreatePnPHost.CreateGraphClient(context);
                    Microsoft.Graph.User graphUser = await graphServiceClient.Users["Karol@karolkozlowski.onmicrosoft.com"].Request().GetAsync();
                    Console.WriteLine($"The display name of test user is {graphUser.DisplayName}");
                    #endregion
                }
            }
        }
    }
}
