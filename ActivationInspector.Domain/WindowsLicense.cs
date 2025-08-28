using System;

namespace ActivationInspector.Domain;

/// <summary>
/// Basic DTO representing a Windows license instance.
/// The structure intentionally mirrors the data extracted from the
/// Check-Activation-Status script but is simplified for demonstration purposes.
/// </summary>
public class WindowsLicense
{
    public string? Name { get; set; }
    public string? Description { get; set; }
    public Guid ActivationId { get; set; }
    public string? LicenseStatus { get; set; }
    public int GraceMinutes { get; set; }
    public DateTime? GraceEndsAt { get; set; }
    public string? PartialProductKey { get; set; }
    public string? Channel { get; set; }
    public string? DigitalPid { get; set; }
    public string? DigitalPid2 { get; set; }
    public DateTime? EvaluationEndDate { get; set; }
    public bool PhoneActivatable { get; set; }
    public string[] Messages { get; set; } = Array.Empty<string>();
}
