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
            var credential = new ClientSecretCredential(
                context.Configuration["GraphTenantId"],
                context.Configuration["GraphClientId"],
                context.Configuration["GraphClientSecret"]);
            return new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        });

        services.AddSingleton(_ =>
        {
            var connectionString = context.Configuration["AzureWebJobsStorage"];
            return new TableServiceClient(connectionString);
        });
    })
    .Build();

host.Run();
