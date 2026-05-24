using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        services.AddSingleton(_ =>
        {
            var credential = new DefaultAzureCredential();
            return new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        });
    })
    .Build();

host.Run();
