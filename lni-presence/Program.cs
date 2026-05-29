using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.EntityFrameworkCore;
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
                var credential = new ClientSecretCredential(
                    context.Configuration["GraphTenantId"],
                    context.Configuration["GraphClientId"],
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
