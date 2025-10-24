using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksMCP.Services;

/// <summary>
/// Manages connection to a running SolidWorks instance
/// </summary>
public class SolidWorksConnection : ISolidWorksConnection, IDisposable
{
    private SldWorks? _application;
    private bool _disposed;

    public SldWorks? Application => _application;

    public ModelDoc2? ActiveDocument => _application?.ActiveDoc as ModelDoc2;

    public bool IsConnected => _application != null;

    public async Task InitializeAsync()
    {
        await Task.Run(() =>
        {
            try
            {
                // Try to connect to running SolidWorks instance using ProgID
                var swType = Type.GetTypeFromProgID("SldWorks.Application");
                if (swType == null)
                {
                    throw new InvalidOperationException(
                        "SolidWorks is not installed or COM registration is missing.");
                }

                // Create or get existing instance
                _application = (SldWorks?)Activator.CreateInstance(swType);

                if (_application == null)
                {
                    throw new InvalidOperationException(
                        "Could not connect to SolidWorks. Please ensure SolidWorks 2020 is running.");
                }

                // Make SolidWorks visible (in case it was hidden)
                _application.Visible = true;

                Console.Error.WriteLine($"Connected to SolidWorks {_application.RevisionNumber()}");
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Failed to connect to SolidWorks. Please ensure SolidWorks 2020 is running.", ex);
            }
        });
    }

    public async Task<bool> ReconnectAsync()
    {
        try
        {
            await InitializeAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;

        // Clean up COM object
        if (_application != null)
        {
            Marshal.ReleaseComObject(_application);
            _application = null;
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
