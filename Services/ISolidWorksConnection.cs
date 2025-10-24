using SolidWorks.Interop.sldworks;

namespace SolidWorksMCP.Services;

/// <summary>
/// Interface for managing connection to SolidWorks application
/// </summary>
public interface ISolidWorksConnection
{
    /// <summary>
    /// Gets the connected SolidWorks application instance
    /// </summary>
    SldWorks? Application { get; }

    /// <summary>
    /// Gets the currently active document
    /// </summary>
    ModelDoc2? ActiveDocument { get; }

    /// <summary>
    /// Initializes connection to SolidWorks
    /// </summary>
    Task InitializeAsync();

    /// <summary>
    /// Checks if SolidWorks is connected and running
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Reconnects to SolidWorks if connection is lost
    /// </summary>
    Task<bool> ReconnectAsync();
}
