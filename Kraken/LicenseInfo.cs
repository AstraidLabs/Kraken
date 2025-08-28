using System;

namespace Kraken;

public class LicenseInfo
{
    public string Application { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string ActivationId { get; set; } = string.Empty;
    public string PartialProductKey { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public int GraceMinutes { get; set; }
    public DateTime? EvaluationEndDate { get; set; }
    public string InstallationId { get; set; } = string.Empty;
}
