using System.ComponentModel;
using ModelContextProtocol.Server;
using SolidWorksMCP.Services;

namespace SolidWorksMCP.Tools;

/// <summary>
/// Tier 4: API documentation search and retrieval tools
/// </summary>
[McpServerToolType]
public class ApiDocumentationTool
{
    [McpServerTool]
    [Description("Searches the SolidWorks API documentation for interfaces, methods, and concepts. Returns relevant documentation with descriptions and syntax.")]
    public static async Task<object> SearchSolidWorksApi(
        IApiDocumentationService apiDocs,
        [Description("Search query (interface name, method name, keywords, or concept)")] string query,
        [Description("Maximum number of results to return (default 10)")] int maxResults = 10)
    {
        try
        {
            var results = await apiDocs.SearchAsync(query, maxResults);

            if (results.Count == 0)
            {
                return new
                {
                    success = true,
                    query = query,
                    resultCount = 0,
                    message = "No documentation found for the query. Try different keywords or check the spelling."
                };
            }

            var formattedResults = results.Select(r => new
            {
                interfaceName = r.InterfaceName,
                methodName = r.MethodName,
                description = r.Description,
                syntax = r.Syntax,
                remarks = r.Remarks,
                relevanceScore = r.RelevanceScore,
                filePath = r.FilePath
            }).ToArray();

            return new
            {
                success = true,
                query = query,
                resultCount = results.Count,
                results = formattedResults
            };
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message
            };
        }
    }

    [McpServerTool]
    [Description("Gets detailed documentation for a specific SolidWorks interface including all methods and properties")]
    public static async Task<object> GetInterfaceDocumentation(
        IApiDocumentationService apiDocs,
        [Description("Interface name (e.g., 'IModelDoc2', 'ISketchManager', 'IFeatureManager')")] string interfaceName)
    {
        try
        {
            var result = await apiDocs.GetInterfaceDocumentationAsync(interfaceName);

            if (result == null)
            {
                return new
                {
                    success = false,
                    error = $"No documentation found for interface: {interfaceName}"
                };
            }

            return new
            {
                success = true,
                interfaceName = result.InterfaceName,
                description = result.Description,
                syntax = result.Syntax,
                remarks = result.Remarks,
                filePath = result.FilePath
            };
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message
            };
        }
    }

    [McpServerTool]
    [Description("Gets documentation for a specific method on a SolidWorks interface including parameters and return values")]
    public static async Task<object> GetMethodDocumentation(
        IApiDocumentationService apiDocs,
        [Description("Interface name (e.g., 'ISketchManager')")] string interfaceName,
        [Description("Method name (e.g., 'CreateLine')")] string methodName)
    {
        try
        {
            var result = await apiDocs.GetMethodDocumentationAsync(interfaceName, methodName);

            if (result == null)
            {
                return new
                {
                    success = false,
                    error = $"No documentation found for method: {interfaceName}.{methodName}"
                };
            }

            return new
            {
                success = true,
                interfaceName = result.InterfaceName,
                methodName = result.MethodName,
                description = result.Description,
                syntax = result.Syntax,
                remarks = result.Remarks,
                filePath = result.FilePath
            };
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message
            };
        }
    }

    [McpServerTool]
    [Description("Gets C# code examples from the SolidWorks API documentation for a specific topic or operation")]
    public static async Task<object> GetCodeExamples(
        IApiDocumentationService apiDocs,
        [Description("Search query for code examples (e.g., 'create extrusion', 'add mate', 'sketch rectangle')")] string query,
        [Description("Maximum number of examples to return (default 5)")] int maxResults = 5)
    {
        try
        {
            var examples = await apiDocs.GetCodeExamplesAsync(query, maxResults);

            if (examples.Count == 0)
            {
                return new
                {
                    success = true,
                    query = query,
                    exampleCount = 0,
                    message = "No code examples found for the query. Try different keywords."
                };
            }

            var formattedExamples = examples.Select(ex => new
            {
                title = ex.Title,
                description = ex.Description,
                language = ex.Language,
                code = ex.Code,
                filePath = ex.FilePath
            }).ToArray();

            return new
            {
                success = true,
                query = query,
                exampleCount = examples.Count,
                examples = formattedExamples
            };
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message
            };
        }
    }

    [McpServerTool]
    [Description("Lists common SolidWorks interfaces and their purposes to help discover API capabilities")]
    public static object ListCommonInterfaces()
    {
        var commonInterfaces = new[]
        {
            new { name = "ISldWorks", purpose = "Top-level application interface. Access to documents, settings, and application operations." },
            new { name = "IModelDoc2", purpose = "Base interface for all document types (parts, assemblies, drawings). Document-level operations." },
            new { name = "IModelDocExtension", purpose = "Extended document operations including selection, custom properties, and advanced features." },
            new { name = "IPartDoc", purpose = "Part-specific operations like bodies, materials, and part features." },
            new { name = "IAssemblyDoc", purpose = "Assembly operations including components, mates, and assembly features." },
            new { name = "IDrawingDoc", purpose = "Drawing-specific operations like sheets, views, and annotations." },
            new { name = "ISketchManager", purpose = "Sketch creation and management. Add lines, circles, arcs, and other sketch entities." },
            new { name = "IFeatureManager", purpose = "Create and manage features like extrusions, cuts, fillets, patterns, etc." },
            new { name = "ISelectionMgr", purpose = "Selection management. Select entities, get selected objects, and manipulate selections." },
            new { name = "IFeature", purpose = "Individual feature object. Access feature properties, suppression state, and feature data." },
            new { name = "ISketch", purpose = "Sketch object. Access sketch segments, points, and relations." },
            new { name = "IBody2", purpose = "Solid or surface body. Access faces, edges, and body properties." },
            new { name = "IFace2", purpose = "Face of a body. Access edges, surface properties, and face operations." },
            new { name = "IEdge", purpose = "Edge of a face or body. Access vertices, curves, and edge properties." },
            new { name = "IMate2", purpose = "Mate constraint in assembly. Define relationships between components." },
            new { name = "IComponent2", purpose = "Assembly component. Access component properties, transform, and suppression." },
            new { name = "IConfiguration", purpose = "Configuration object. Manage design variations and configurations." },
            new { name = "IDimension", purpose = "Dimension object. Access and modify dimension values." },
            new { name = "IDisplayDimension", purpose = "Display dimension properties and appearance." }
        };

        return new
        {
            success = true,
            interfaceCount = commonInterfaces.Length,
            interfaces = commonInterfaces,
            note = "Use SearchSolidWorksApi or GetInterfaceDocumentation tools to get detailed information about any interface."
        };
    }

    [McpServerTool]
    [Description("Provides workflow guidance for common SolidWorks automation tasks")]
    public static object GetWorkflowGuidance(
        [Description("Type of workflow: 'part', 'assembly', 'drawing', 'sketch', 'feature', 'selection', or 'custom'")] string workflowType)
    {
        var workflows = new Dictionary<string, object>
        {
            ["part"] = new
            {
                workflow = "Creating a Part",
                steps = new[]
                {
                    "1. Create new part: ISldWorks.NewDocument() or use CreatePart tool",
                    "2. Select a plane: IModelDocExtension.SelectByID2(\"Front Plane\", \"PLANE\", ...)",
                    "3. Create sketch: ISketchManager.InsertSketch() or use CreateSketch tool",
                    "4. Add sketch geometry: ISketchManager.CreateLine(), CreateCircle(), etc.",
                    "5. Exit sketch: ISketchManager.InsertSketch(true)",
                    "6. Create feature: IFeatureManager.FeatureExtrusion2() or use CreateExtrude tool",
                    "7. Save document: IModelDoc2.Save3()"
                },
                commonInterfaces = new[] { "ISldWorks", "IModelDoc2", "ISketchManager", "IFeatureManager" }
            },
            ["assembly"] = new
            {
                workflow = "Creating an Assembly",
                steps = new[]
                {
                    "1. Create new assembly: ISldWorks.NewDocument() or use CreateAssembly tool",
                    "2. Add components: IAssemblyDoc.AddComponent5()",
                    "3. Position components: IComponent2.Transform2",
                    "4. Add mates: IAssemblyDoc.AddMate5()",
                    "5. Save assembly: IModelDoc2.Save3()"
                },
                commonInterfaces = new[] { "ISldWorks", "IAssemblyDoc", "IComponent2", "IMate2" }
            },
            ["sketch"] = new
            {
                workflow = "Working with Sketches",
                steps = new[]
                {
                    "1. Get sketch manager: IModelDoc2.SketchManager",
                    "2. Select plane or face: IModelDocExtension.SelectByID2()",
                    "3. Insert sketch: ISketchManager.InsertSketch()",
                    "4. Create geometry: CreateLine(), CreateCircle(), CreateRectangle(), etc.",
                    "5. Add constraints: ISketch.AddConstraints()",
                    "6. Add dimensions: ISketchManager.CreateDimension()",
                    "7. Exit sketch: ISketchManager.InsertSketch(true)"
                },
                commonInterfaces = new[] { "ISketchManager", "ISketch", "ISketchSegment" }
            },
            ["feature"] = new
            {
                workflow = "Creating Features",
                steps = new[]
                {
                    "1. Ensure sketch or geometry is selected",
                    "2. Get feature manager: IModelDoc2.FeatureManager",
                    "3. Call feature method: FeatureExtrusion2(), FeatureCut4(), FeatureFillet3(), etc.",
                    "4. Access feature data if needed: IFeature.GetDefinition()",
                    "5. Modify and update: IFeatureDefinition.AccessSelections(), ReleaseSelectionAccess()"
                },
                commonInterfaces = new[] { "IFeatureManager", "IFeature", "IExtrudeFeatureData2", "IFilletFeatureData" }
            },
            ["selection"] = new
            {
                workflow = "Selection and Object Access",
                steps = new[]
                {
                    "1. Select by ID: IModelDocExtension.SelectByID2(name, type, x, y, z, ...)",
                    "2. Or select by ray: IModelDocExtension.SelectByRay()",
                    "3. Get selection manager: IModelDoc2.SelectionManager",
                    "4. Get selected objects: ISelectionMgr.GetSelectedObject6()",
                    "5. Get selection type: ISelectionMgr.GetSelectedObjectType3()",
                    "6. Process selected objects",
                    "7. Clear selection: IModelDoc2.ClearSelection2()"
                },
                commonInterfaces = new[] { "IModelDocExtension", "ISelectionMgr" }
            },
            ["drawing"] = new
            {
                workflow = "Creating a Drawing",
                steps = new[]
                {
                    "1. Create drawing: ISldWorks.NewDocument()",
                    "2. Get drawing doc: IDrawingDoc",
                    "3. Create sheet: IDrawingDoc.CreateDrawViewFromModelView3()",
                    "4. Add views: IDrawingDoc.CreateDrawViewFromModelView3()",
                    "5. Add annotations: IDrawingDoc.InsertNote(), InsertDimension()",
                    "6. Save drawing: IModelDoc2.Save3()"
                },
                commonInterfaces = new[] { "IDrawingDoc", "IView", "ISheet", "INote", "IDisplayDimension" }
            }
        };

        var key = workflowType.ToLowerInvariant();
        if (workflows.ContainsKey(key))
        {
            return new
            {
                success = true,
                workflowType = workflowType,
                guidance = workflows[key]
            };
        }
        else
        {
            return new
            {
                success = false,
                error = $"Unknown workflow type: {workflowType}",
                availableTypes = workflows.Keys.ToArray()
            };
        }
    }
}
