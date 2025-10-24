using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksMCP.Services;

namespace SolidWorksMCP.Resources;

/// <summary>
/// Tier 3: Simplified resources for exposing SolidWorks state as readable data
/// </summary>
[McpServerResourceType]
public class SolidWorksResourcesSimple
{
    [McpServerResource(
        UriTemplate = "solidworks://active-document",
        Name = "Active Document",
        MimeType = "application/json")]
    [Description("Gets information about the currently active document including type, path, and modification status")]
    public static object GetActiveDocumentResource(ISolidWorksConnection connection)
    {
        try
        {
            if (!connection.IsConnected)
                return new { uri = "solidworks://active-document", error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { uri = "solidworks://active-document", error = "No active document" };

            var docType = (swDocumentTypes_e)doc.GetType();

            return new
            {
                uri = "solidworks://active-document",
                content = new
                {
                    name = doc.GetTitle(),
                    type = docType.ToString(),
                    path = doc.GetPathName(),
                    isModified = doc.GetSaveFlag(),
                    isReadOnly = doc.IsOpenedReadOnly(),
                    visible = doc.Visible
                }
            };
        }
        catch (Exception ex)
        {
            return new { uri = "solidworks://active-document", error = ex.Message };
        }
    }

    [McpServerResource(
        UriTemplate = "solidworks://feature-tree",
        Name = "Feature Tree",
        MimeType = "application/json")]
    [Description("Gets the feature tree of the active document, showing all features")]
    public static object GetFeatureTreeResource(ISolidWorksConnection connection)
    {
        try
        {
            if (!connection.IsConnected)
                return new { uri = "solidworks://feature-tree", error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { uri = "solidworks://feature-tree", error = "No active document" };

            var features = new List<object>();
            var feature = (Feature?)doc.FirstFeature();

            while (feature != null)
            {
                features.Add(new
                {
                    name = feature.Name,
                    typeName = feature.GetTypeName(),
                    visible = feature.Visible,
                    suppressed = feature.IsSuppressed()
                });
                feature = (Feature?)feature.GetNextFeature();
            }

            return new
            {
                uri = "solidworks://feature-tree",
                content = new
                {
                    documentName = doc.GetTitle(),
                    featureCount = features.Count,
                    features = features
                }
            };
        }
        catch (Exception ex)
        {
            return new { uri = "solidworks://feature-tree", error = ex.Message };
        }
    }
}
