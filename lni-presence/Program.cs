using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddSingleton(sp =>
        {
            var logger = sp.GetRequiredService<ILoggerFactory>().CreateLogger("GraphClientSetup");
            var debugMode = context.Configuration["DEBUG_SELF_PRESENCE"];
            if (string.Equals(debugMode, "true", StringComparison.OrdinalIgnoreCase))
            {
                var tenantId = context.Configuration["AZURE_TENANT_ID"];
                var clientId = context.Configuration["AZURE_CLIENT_ID"];
                logger.LogInformation("Graph client using InteractiveBrowserCredential with ClientId: {ClientId}, TenantId: {TenantId}", clientId, tenantId);
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
                var tenantId = context.Configuration["GraphTenantId"];
                var clientId = context.Configuration["GraphClientId"];
                logger.LogInformation("Graph client using ClientSecretCredential with ClientId: {ClientId}, TenantId: {TenantId}", clientId, tenantId);
                var credential = new ClientSecretCredential(
                    tenantId,
                    clientId,
                    context.Configuration["GraphClientSecret"]);
                return new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
            }
        });

        var sqlConnectionString = context.Configuration["SqlConnectionString"];
        services.AddDbContext<lni_presence.PresenceDbContext>(options =>
            options.UseSqlServer(sqlConnectionString));
    })
    .Build();

host.Run();
