// Creates and configures the PnP host
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Services;
using PnPCore_.NET_7._0.PnPHost;

const string SITETARGET = "site";

var host = new CreatePnPHost();
await host.PnPIHost.StartAsync();
using (var scope = host.PnPIHost.Services.CreateScope())
{
    // Create the PnPContext
    // Documentation -> https://pnp.github.io/pnpcore/
    using (var context = await scope.ServiceProvider.GetRequiredService<IPnPContextFactory>().CreateAsync(name: SITETARGET))
    {
        // Load the Title property of the site's root web
        await context.Web.LoadAsync(p => p.Title);
        Console.WriteLine($"The title of the web is {context.Web.Title}");

        #region CSOM Context
        // Create a CSOM Context
        // Documentation -> https://pnp.github.io/pnpframework/
        using (Microsoft.SharePoint.Client.ClientContext csomContext = CreatePnPHost.GetCSOMContext(context))
        {
            Microsoft.SharePoint.Client.Web web = csomContext.Web;
            csomContext.Load(web);
            csomContext.ExecuteQuery();
            Console.WriteLine($"The title of the web is {web.Title}");
        }
        #endregion

        #region Graph SDK Context
        // Create a Graph Context
        // Documentation -> https://learn.microsoft.com/pl-pl/graph/api/overview?view=graph-rest-1.0
        Microsoft.Graph.GraphServiceClient graphServiceClient = CreatePnPHost.GetGraphContext(context);
        Microsoft.Graph.Models.User graphUser = await graphServiceClient.Users["user@contoso.onmicrosoft.com"].GetAsync();
        Console.WriteLine($"The display name of test user is {graphUser.DisplayName}");
        #endregion
    }
}
