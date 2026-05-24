using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Communications.GetPresencesByUserId;
using Microsoft.Graph.Models;

namespace lni_presence
{
    public class PresencePoller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly string _securityGroupId;
        private readonly ILogger<PresencePoller> _logger;

        public PresencePoller(GraphServiceClient graphClient, IConfiguration configuration, ILogger<PresencePoller> logger)
        {
            _graphClient = graphClient;
            _securityGroupId = configuration["SecurityGroupId"]
                ?? throw new InvalidOperationException("SecurityGroupId is not configured.");
            _logger = logger;
        }

        [Function("PresencePoller")]
        public async Task Run([TimerTrigger("0 */1 * * * *")] TimerInfo timerInfo)
        {
            _logger.LogInformation("Presence polling started at {Time}", DateTime.UtcNow);

            var userIds = new List<string>();
            var userNames = new Dictionary<string, string>();

            // Get group members with pagination
            var membersResponse = await _graphClient.Groups[_securityGroupId].Members
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Select = ["id", "displayName"];
                });

            while (membersResponse?.Value != null)
            {
                foreach (var member in membersResponse.Value)
                {
                    if (member is User user && user.Id != null)
                    {
                        userIds.Add(user.Id);
                        userNames[user.Id] = user.DisplayName ?? "Unknown";
                    }
                }

                // Follow pagination
                if (membersResponse.OdataNextLink != null)
                {
                    membersResponse = await _graphClient.Groups[_securityGroupId].Members
                        .GetAsync(config =>
                        {
                            config.QueryParameters.Top = 100;
                            config.QueryParameters.Skip = userIds.Count;
                        });
                }
                else
                {
                    break;
                }
            }

            _logger.LogInformation("Found {Count} users in security group", userIds.Count);

            if (userIds.Count == 0)
                return;

            // Get presence in batches of 650 (API limit)
            const int batchSize = 650;
            for (int i = 0; i < userIds.Count; i += batchSize)
            {
                var batch = userIds.Skip(i).Take(batchSize).ToList();

                var requestBody = new GetPresencesByUserIdPostRequestBody
                {
                    Ids = batch
                };

                var presenceResponse = await _graphClient.Communications.GetPresencesByUserId
                    .PostAsGetPresencesByUserIdPostResponseAsync(requestBody);

                if (presenceResponse?.Value != null)
                {
                    foreach (var presence in presenceResponse.Value)
                    {
                        var displayName = presence.Id != null && userNames.TryGetValue(presence.Id, out var name)
                            ? name
                            : "Unknown";

                        _logger.LogInformation(
                            "User: {DisplayName} | Availability: {Availability} | Activity: {Activity}",
                            displayName,
                            presence.Availability ?? "N/A",
                            presence.Activity ?? "N/A");
                    }
                }
            }

            _logger.LogInformation("Presence polling completed. Polled {Count} users.", userIds.Count);
        }
    }
}
