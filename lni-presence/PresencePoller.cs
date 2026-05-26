using Microsoft.Azure.Functions.Worker;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Communications.GetPresencesByUserId;
using Microsoft.Graph.Models;

namespace lni_presence;

public class PresencePoller
{
    private readonly GraphServiceClient _graphClient;
    private readonly PresenceDbContext _db;
    private readonly string _securityGroupId;
    private readonly ILogger<PresencePoller> _logger;

    public PresencePoller(
        GraphServiceClient graphClient,
        PresenceDbContext db,
        IConfiguration configuration,
        ILogger<PresencePoller> logger)
    {
        _graphClient = graphClient;
        _db = db;
        _securityGroupId = configuration["SecurityGroupId"] ?? "";
        _logger = logger;
    }

    [Function("PresencePoller")]
    public async Task Run([TimerTrigger("0 0 0 * * *")] TimerInfo timerInfo)
    {
        await _db.Database.EnsureCreatedAsync();

        var debugMode = Environment.GetEnvironmentVariable("DEBUG_SELF_PRESENCE");
        if (string.Equals(debugMode, "true", StringComparison.OrdinalIgnoreCase))
        {
            await RunDebugSelfPresence();
            return;
        }

        await RunProductionPolling();
    }

    private async Task RunDebugSelfPresence()
    {
        _logger.LogInformation("DEBUG: Polling own presence at {Time}", DateTime.UtcNow);

        var me = await _graphClient.Me.GetAsync(config =>
        {
            config.QueryParameters.Select = ["id", "displayName"];
        });

        if (me?.Id == null)
        {
            _logger.LogError("DEBUG: Could not retrieve current user");
            return;
        }

        var presence = await _graphClient.Me.Presence.GetAsync();
        if (presence == null)
        {
            _logger.LogError("DEBUG: Could not retrieve presence for current user");
            return;
        }

        var userId = me.Id;
        var displayName = me.DisplayName ?? "Unknown";
        var availability = presence.Availability ?? "Unknown";
        var activity = presence.Activity ?? "Unknown";
        var now = DateTimeOffset.UtcNow;

        _logger.LogInformation(
            "DEBUG: {DisplayName} | Availability: {Availability} | Activity: {Activity}",
            displayName, availability, activity);

        var existing = await _db.PresenceLog
            .FirstOrDefaultAsync(r => r.UserId == userId && r.IsCurrent);

        if (existing != null && existing.Availability == availability && existing.Activity == activity)
        {
            _logger.LogInformation("DEBUG: No status change detected");
            return;
        }

        if (existing != null)
        {
            existing.IsCurrent = false;
            existing.EndTime = now;

            _logger.LogInformation(
                "DEBUG: Status changed: {OldAvail} -> {NewAvail} | {OldActivity} -> {NewActivity}",
                existing.Availability, availability, existing.Activity, activity);
        }

        _db.PresenceLog.Add(new PresenceRecord
        {
            UserId = userId,
            DisplayName = displayName,
            Availability = availability,
            Activity = activity,
            StartTime = now,
            EndTime = null,
            IsCurrent = true
        });

        await _db.SaveChangesAsync();
        _logger.LogInformation("DEBUG: Presence record saved");
    }

    private async Task RunProductionPolling()
    {
        _logger.LogInformation("Presence polling started at {Time}", DateTime.UtcNow);

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

        // Step 2: Load all current presence records from DB
        var currentRecords = await _db.PresenceLog
            .Where(r => r.IsCurrent)
            .ToDictionaryAsync(r => r.UserId);

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
                if (existing.Availability == availability && existing.Activity == activity)
                    continue;

                existing.IsCurrent = false;
                existing.EndTime = now;

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

            _db.PresenceLog.Add(new PresenceRecord
            {
                UserId = userId,
                DisplayName = displayName,
                Availability = availability,
                Activity = activity,
                StartTime = now,
                EndTime = null,
                IsCurrent = true
            });
            changesDetected++;
        }

        await _db.SaveChangesAsync();

        _logger.LogInformation(
            "Presence polling completed. Polled {Total} users, {Changes} changes detected.",
            allPresence.Count, changesDetected);
    }
}
