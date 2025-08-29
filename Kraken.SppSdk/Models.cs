namespace Kraken.SppSdk;

public enum LicenseState
{
    Unlicensed = 0,
    Licensed = 1,
    OutOfBoxGrace = 2,
    OutOfToleranceGrace = 3,
    NonGenuineGrace = 4,
    Notification = 5,
    ExtendedGrace = 6
}

public record WindowsLicenseInfo(Guid Slid, string ProductKey, DateTime? Expiry, LicenseState State);
public record OfficeLicenseInfo(Guid Slid, string ProductKey, DateTime? Expiry, LicenseState State, string Edition);
public record VNextLicense(string FileName, string ProductReleaseId, string Status, DateTime? Expiry);
public record SubStatus(int LicenseStatus, int LicenseState, int GenuineStatus, int GenuineState);
public record SppLicenseStatus(uint Status, uint GraceMinutes, uint ReasonHResult, ulong ValidityFileTimeUtc);
