using ActivationInspector.Domain;
using ActivationInspector.Infrastructure.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ActivationInspector.Infrastructure.Licensing;

/// <summary>
/// Provides a minimal wrapper around the native Software Licensing API used by Windows.
/// The real script exposes a vast amount of information; here we only query whether the
/// current installation is considered genuine by the local licensing service.
/// </summary>
public class WindowsLicensingProvider
{
    public Task<IReadOnlyList<WindowsLicense>> GetLicensesAsync(CancellationToken token)
    {
        return Task.Run(() =>
        {
            var list = new List<WindowsLicense>();
            try
            {
                uint genuine = 0;
                int hr = SppNative.SLIsWindowsGenuineLocal(ref genuine);
                var dto = new WindowsLicense
                {
                    Name = "Windows",
                    LicenseStatus = hr == 0 ? (genuine == 1 ? "Licensed" : "Unlicensed") : $"HRESULT 0x{hr:X8}"
                };
                list.Add(dto);
            }
            catch (DllNotFoundException)
            {
                list.Add(new WindowsLicense
                {
                    Name = "Windows",
                    LicenseStatus = "Licensing APIs not available"
                });
            }
            catch (Exception ex)
            {
                list.Add(new WindowsLicense
                {
                    Name = "Windows",
                    LicenseStatus = ex.Message
                });
            }
            return (IReadOnlyList<WindowsLicense>)list;
        }, token);
    }
}
