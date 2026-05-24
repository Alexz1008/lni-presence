using Azure.Data.Tables;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddSingleton(_ =>
        {
            var debugMode = context.Configuration["DEBUG_SELF_PRESENCE"];
            if (string.Equals(debugMode, "true", StringComparison.OrdinalIgnoreCase))
            {
                var tenantId = context.Configuration["AZURE_TENANT_ID"];
                var clientId = context.Configuration["AZURE_CLIENT_ID"];
                var credential = new InteractiveBrowserCredential(
                    new InteractiveBrowserCredentialOptions
                    {
                        TenantId = tenantId,
                        ClientId = clientId,
                        RedirectUri = new Uri("http://localhost:1234")
                    });
                return new GraphServiceClient(credential, ["Presence.Read", "User.Read"]);
            }
            else
            {
                var credential = new DefaultAzureCredential();
                return new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
            }
        });

        services.AddSingleton(_ =>
        {
            var connectionString = context.Configuration["AzureWebJobsStorage"];
            return new TableServiceClient(connectionString);
        });
    })
    .Build();

host.Run();
