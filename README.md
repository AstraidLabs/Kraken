# Kraken

WPF application that queries Windows, Office and related licence information using the Software Protection Platform (SPP) APIs.

## Prerequisites

- Windows 10/11
- .NET 8 SDK or later
- Administrator privileges may be required to read some registry keys

## Build

```
dotnet build Kraken/Kraken.csproj
```

## Run

```
dotnet run --project Kraken/Kraken.csproj
```

The application displays licence data and allows exporting the summary to JSON using the **Save JSON** button.
