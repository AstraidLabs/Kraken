namespace Kraken;

/// <summary>
/// Helper extension methods for working with <see cref="SppApi.SppValue"/> values.
/// </summary>
public static class SppExtension
{
    public static string? AsString(this SppApi.SppValue value) => value.S;
    public static uint? AsUInt32(this SppApi.SppValue value) => value.U32;
    public static ulong? AsUInt64(this SppApi.SppValue value) => value.U64;
}

