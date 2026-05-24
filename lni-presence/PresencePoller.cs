using Azure.Data.Tables;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Communications.GetPresencesByUserId;
using Microsoft.Graph.Models;

namespace lni_presence;

public class PresencePoller
{
    private readonly GraphServiceClient _graphClient;
    private readonly TableServiceClient _tableServiceClient;
    private readonly string _securityGroupId;
    private readonly ILogger<PresencePoller> _logger;
    private const string TableName = "PresenceLog";

    public PresencePoller(
        GraphServiceClient graphClient,
        TableServiceClient tableServiceClient,
        IConfiguration configuration,
        ILogger<PresencePoller> logger)
    {
        _graphClient = graphClient;
        _tableServiceClient = tableServiceClient;
        _securityGroupId = configuration["SecurityGroupId"]
            ?? throw new InvalidOperationException("SecurityGroupId is not configured.");
        _logger = logger;
    }

    [Function("PresencePoller")]
    public async Task Run([TimerTrigger("0 0 0 * * *")] TimerInfo timerInfo)
    {
        _logger.LogInformation("Presence polling started at {Time}", DateTime.UtcNow);

        var tableClient = _tableServiceClient.GetTableClient(TableName);
        await tableClient.CreateIfNotExistsAsync();

        // Step 1: Get group members with pagination
        var userIds = new List<string>();
        var userNames = new Dictionary<string, string>();

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
        if (userIds.Count == 0) return;

        // Step 2: Load all current presence records from table
        var currentRecords = new Dictionary<string, PresenceRecord>();
        await foreach (var entity in tableClient.QueryAsync<PresenceRecord>(
            filter: $"PartitionKey eq 'CURRENT'"))
        {
            currentRecords[entity.RowKey] = entity;
        }

        // Step 3: Poll presence in batches of 650 (API limit)
        var allPresence = new List<Presence>();
        const int batchSize = 650;
        for (int i = 0; i < userIds.Count; i += batchSize)
        {
            var batch = userIds.Skip(i).Take(batchSize).ToList();
            var requestBody = new GetPresencesByUserIdPostRequestBody { Ids = batch };
            var presenceResponse = await _graphClient.Communications.GetPresencesByUserId
                .PostAsGetPresencesByUserIdPostResponseAsync(requestBody);

            if (presenceResponse?.Value != null)
                allPresence.AddRange(presenceResponse.Value);
        }

        // Step 4: Compare with previous state and write changes
        var now = DateTimeOffset.UtcNow;
        int changesDetected = 0;

        foreach (var presence in allPresence)
        {
            if (presence.Id == null) continue;

            var userId = presence.Id;
            var availability = presence.Availability ?? "Unknown";
            var activity = presence.Activity ?? "Unknown";
            var displayName = userNames.GetValueOrDefault(userId, "Unknown");

            if (currentRecords.TryGetValue(userId, out var existing))
            {
                // Status unchanged — skip
                if (existing.Availability == availability && existing.Activity == activity)
                    continue;

                // Status changed — archive previous as history row with EndTime
                var historyRowKey = (DateTimeOffset.MaxValue.Ticks - existing.StartTime.Ticks)
                    .ToString("D19");
                var historyRecord = new PresenceRecord
                {
                    PartitionKey = userId,
                    RowKey = historyRowKey,
                    UserId = userId,
                    DisplayName = existing.DisplayName,
                    Availability = existing.Availability,
                    Activity = existing.Activity,
                    StartTime = existing.StartTime,
                    EndTime = now
                };
                await tableClient.UpsertEntityAsync(historyRecord);

                _logger.LogInformation(
                    "Status changed: {DisplayName} | {OldAvail} -> {NewAvail} | {OldActivity} -> {NewActivity}",
                    displayName, existing.Availability, availability, existing.Activity, activity);
            }
            else
            {
                _logger.LogInformation(
                    "New user tracked: {DisplayName} | {Availability} | {Activity}",
                    displayName, availability, activity);
            }

            // Upsert new current status with StartTime = now
            var currentRecord = new PresenceRecord
            {
                PartitionKey = "CURRENT",
                RowKey = userId,
                UserId = userId,
                DisplayName = displayName,
                Availability = availability,
                Activity = activity,
                StartTime = now,
                EndTime = null
            };
            await tableClient.UpsertEntityAsync(currentRecord);
            changesDetected++;
        }

        _logger.LogInformation(
            "Presence polling completed. Polled {Total} users, {Changes} changes detected.",
            allPresence.Count, changesDetected);
    }
}
