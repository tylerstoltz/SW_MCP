# SolidWorks MCP Server

A comprehensive Model Context Protocol (MCP) server for SolidWorks 2020, enabling natural language control of SolidWorks through Claude Desktop and other MCP clients.

## Architecture

This server implements a **4-tool types** that combines high-level workflow tools with flexible, dynamic API execution:

### 1: High-Level Workflow Tools
Simple, deterministic tools for common operations:
- `create_part` - Create new part documents
- `create_assembly` - Create new assembly documents
- `create_sketch` - Create sketches on planes
- `create_extrude` - Create extrusion features
- `create_rectangle` - Add rectangles to sketches
- `create_circle` - Add circles to sketches
- `create_line` - Add lines to sketches
- `create_base_flange` - Create sheet metal base flange features (first feature in sheet metal parts)
- `create_edge_flange` - Create edge flanges on sheet metal parts
- `get_model_info` - Get document information
- `save_document` - Save documents

### 2: Dynamic API Execution
Flexible tools for arbitrary SolidWorks API calls:
- `execute_solidworks_api` - Execute any SolidWorks API method
- `access_solidworks_property` - Get/set properties on interfaces

### 3: Resources for State Inspection
Expose SolidWorks state as readable data:
- `get_active_document_resource` - Current document info
- `get_feature_tree_resource` - Feature list and hierarchy

**Note:** Additional resource methods (selection, sketches, configurations) are available in `SolidWorksResources.cs.bak` but require API signature adjustments for SolidWorks 2020.

### 4: API Documentation Search
Self-documenting capabilities:
- `search_solidworks_api` - Full-text search of API documentation
- `get_interface_documentation` - Get interface details
- `get_method_documentation` - Get method details
- `get_code_examples` - Find C# code examples
- `list_common_interfaces` - List common SolidWorks interfaces
- `get_workflow_guidance` - Get workflow guidance for common tasks

## Prerequisites

1. **SolidWorks 2020** installed (does not need to be running - the server will connect/start it)
2. **.NET 8.0 SDK** or later
3. **Windows OS** (required for COM interop)
4. **SolidWorks API DLLs** in the correct location:
   - Default: `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\`
   - Files: `SolidWorks.Interop.sldworks.dll`, `SolidWorks.Interop.swconst.dll`
5. **SolidWorks API Documentation** (optional but recommended for search features):
   - Path: `Path\to\solidworks\Solidworks-SDK\docs\api\` - convert from CHM if needed
   - Be sure to update this in "ApiDocumentationService.cs": `_apiDocsPath = @"Path\\to\\solidworks\\Solidworks-SDK\\docs\\api\\";` is placeholder (add to config later)

## Installation

### 1. Build the Project

```bash
cd "Path\\To\\SW_MCP"
dotnet restore
dotnet build -c Release
```

### 2. Verify SolidWorks API DLL Paths

If your SolidWorks installation is in a different location, update the paths in `SolidWorksMCP.csproj`:

```xml
<Reference Include="SolidWorks.Interop.sldworks">
  <HintPath>YOUR_PATH_HERE\SolidWorks.Interop.sldworks.dll</HintPath>
  <EmbedInteropTypes>false</EmbedInteropTypes>
</Reference>
```

### 3. Configure Claude Desktop

Add the following to your Claude Desktop configuration file:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "solidworks": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "Path\\To\\SW_MCP\\SolidWorksMCP.csproj"
      ]
    }
  }
}
```

Or use the built executable:

```json
{
  "mcpServers": {
    "solidworks": {
      "command": "Path\\To\\SW_MCP\\bin\\Release\\net8.0\\SolidWorksMCP.exe"
    }
  }
}
```

### 4. Restart Claude Desktop

Restart Claude Desktop to load the new MCP server.

## Usage

### Basic Workflow Example

```
User: "Create a new part and draw a 100mm x 50mm rectangle, then extrude it 25mm"

Claude will:
1. Use create_part to create a new part
2. Use create_sketch to start a sketch on the Front plane
3. Use create_rectangle with coordinates (0, 0) to (0.1, 0.05) [meters]
4. Use create_extrude with depth 0.025 [meters]
```

### Sheet Metal Example

```
User: "Create a sheet metal part with a 100mm x 80mm base flange, 2mm thick"

Claude will:
1. Use create_part to create a new part
2. Use create_sketch to start a sketch on the Front plane
3. Use create_rectangle with coordinates (0, 0) to (0.1, 0.08) [meters]
4. Use create_base_flange with thickness 0.002 [meters]
```

### Advanced API Calls

```
User: "Use the SolidWorks API to create a fillet with radius 5mm on the selected edges"

Claude will:
1. Use search_solidworks_api to find fillet-related methods
2. Use execute_solidworks_api to call IFeatureManager.FeatureFillet with appropriate parameters
```

### Inspecting State

```
User: "What features are in the current model?"

Claude will use get_feature_tree_resource to list all features.
```

### Learning the API

```
User: "How do I create a mate in an assembly?"

Claude will:
1. Use get_workflow_guidance with workflowType: "assembly"
2. Or search_solidworks_api for "mate"
3. Or get_code_examples for "add mate"
```

## Units

**IMPORTANT:** All dimensions in the SolidWorks API are in **meters**, regardless of the document units.

- 1 mm = 0.001 meters
- 1 inch = 0.0254 meters

The workflow tools expect dimensions in meters. Convert your dimensions accordingly.

## Architecture Details

### Connection Management

The server connects to SolidWorks using COM interop via `Type.GetTypeFromProgID("SldWorks.Application")` and `Activator.CreateInstance()`. This will either connect to a running instance or start SolidWorks if it's not running. The connection is established on server startup and maintained throughout the session.

### Error Handling

All tools return structured responses with `success` boolean and `error` messages:

```json
{
  "success": true,
  "message": "Operation completed successfully"
}
```

Or on failure:

```json
{
  "success": false,
  "error": "Error message describing what went wrong"
}
```

### API Documentation Search

The documentation search service indexes the HTML documentation files on first use and caches the results for the session. It uses:
- Full-text search across 13,000+ HTML files
- Relevance scoring based on term frequency and file names
- Extraction of descriptions, syntax, and remarks from structured documentation

## Troubleshooting

### "Not connected to SolidWorks" Error

**Solution:** Ensure SolidWorks is installed properly. The server will attempt to start it automatically.

### "Could not connect to SolidWorks" on Startup

**Solutions:**
1. Verify SolidWorks 2020 is installed (check COM registration)
2. Check that the SolidWorks API DLLs are in the correct location
3. Try running Claude Desktop as Administrator
4. Ensure SolidWorks license is valid and activated
5. Check Windows Event Viewer for COM-related errors

### "No active document" Error

**Solution:** Open or create a document in SolidWorks before performing document-specific operations.

### API Documentation Not Found

**Solution:** Update the documentation path in `ApiDocumentationService.cs`:

```csharp
_apiDocsPath = @"YOUR_DOCS_PATH_HERE";
```

### Build Errors Related to SolidWorks Interop DLLs

**Solutions:**
1. Verify DLL paths in `.csproj` file
2. Check that SolidWorks 2020 is installed (not just viewer)
3. Ensure you have access to the `api\redist` folder

## Development

### Project Structure

```
SolidWorksMCP/
├── Program.cs                          # Entry point and DI configuration
├── SolidWorksMCP.csproj               # Project file with dependencies
├── Services/
│   ├── ISolidWorksConnection.cs       # Connection interface
│   ├── SolidWorksConnection.cs        # COM connection implementation
│   ├── IApiDocumentationService.cs    # Documentation search interface
│   └── ApiDocumentationService.cs     # Documentation search implementation
├── Tools/
│   ├── WorkflowTools.cs               # Tier 1: High-level workflow tools
│   ├── DynamicApiTool.cs              # Tier 2: Dynamic API execution
│   └── ApiDocumentationTool.cs        # Tier 4: Documentation search tools
└── Resources/
    └── SolidWorksResources.cs         # Tier 3: State inspection resources
```

### Adding New Tools

1. Create a new static class in the `Tools/` folder
2. Decorate the class with `[McpServerToolType]`
3. Add static methods decorated with `[McpServerTool]`
4. Use `[Description(...)]` attributes for parameter documentation
5. Inject services like `ISolidWorksConnection` as method parameters

Example:

```csharp
[McpServerToolType]
public class MyCustomTools
{
    [McpServerTool]
    [Description("Does something useful")]
    public static object MyTool(
        ISolidWorksConnection connection,
        [Description("A parameter")] string param1)
    {
        // Implementation
        return new { success = true, result = "..." };
    }
}
```

### Extending Interface Mappings

To support more interface types in `execute_solidworks_api`, add mappings in `DynamicApiTool.GetInterfaceInstance()`:

```csharp
return interfaceName.ToLowerInvariant() switch
{
    // ... existing mappings ...
    "imynewinterface" => GetMyNewInterface(connection),
    _ => null
};
```

## Limitations

1. **Single Instance:** Connects to one running SolidWorks instance
2. **No Macro Recording:** Does not record operations as macros
3. **Synchronous Execution:** Operations block until complete
4. **COM Limitations:** Subject to COM threading and marshaling constraints
5. **Read-Only COM Objects:** Some COM objects cannot be fully serialized in responses

## License

This project is provided as-is for use with SolidWorks 2020. Ensure compliance with SolidWorks API licensing terms.

## References

- [Model Context Protocol](https://modelcontextprotocol.io/)
- [MCP C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)
- [SolidWorks API Documentation](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks_namespace.html)
