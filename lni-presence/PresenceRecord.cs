using System.ComponentModel.DataAnnotations;

namespace lni_presence;

public class PresenceRecord
{
    public long Id { get; set; }

    [MaxLength(200)]
    public string UserId { get; set; } = string.Empty;

    [MaxLength(200)]
    public string DisplayName { get; set; } = string.Empty;

    [MaxLength(50)]
    public string Availability { get; set; } = string.Empty;

    [MaxLength(50)]
    public string Activity { get; set; } = string.Empty;

    public DateTimeOffset StartTime { get; set; }
    public DateTimeOffset? EndTime { get; set; }

    public bool IsCurrent { get; set; }
}
