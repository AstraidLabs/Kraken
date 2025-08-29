using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace Kraken;

/// <summary>
/// Dependency injection helpers for registering Software Protection Platform services.
/// </summary>
public static class SppDI
{
    /// <summary>
    /// Adds SPP session support to the service collection.
    /// </summary>
    public static IServiceCollection AddSpp(this IServiceCollection services)
    {
        services.AddSingleton<SppSession>(_ =>
        {
            try
            {
                return SppSession.Open();
            }
            catch (SppException ex)
            {
                Log.Error(ex, "Unable to open SPP session");
                throw;
            }
        });
        return services;
    }
}

