namespace Kraken.SppSdk;

public sealed class SppException : Exception
{
    public int HResultCode { get; }
    public string? Function { get; }

    public SppException(int hresult, string? function = null) : base($"SPP call failed with HRESULT 0x{hresult:X8}" + (function != null ? $" in {function}" : string.Empty))
    {
        HResultCode = hresult;
        Function = function;
    }
}
