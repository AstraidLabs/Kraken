using System;
using Serilog;

namespace Kraken;

/// <summary>
/// Represents errors that occur when interacting with the Software Protection Platform.
/// Logs the error using Serilog upon creation.
/// </summary>
public class SppException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SppException"/> class with a specified error message.
    /// </summary>
    public SppException(string message) : base(message)
    {
        Log.Error(message);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SppException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
    /// </summary>
    public SppException(string message, Exception innerException) : base(message, innerException)
    {
        Log.Error(innerException, message);
    }
}

