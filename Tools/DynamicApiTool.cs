using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using ModelContextProtocol.Server;
using SolidWorks.Interop.sldworks;
using SolidWorksMCP.Services;

namespace SolidWorksMCP.Tools;

/// <summary>
/// Tier 2: Dynamic API execution tool for flexible, arbitrary SolidWorks API calls
/// </summary>
[McpServerToolType]
public class DynamicApiTool
{
    [McpServerTool]
    [Description("Executes an arbitrary SolidWorks API method call. Use this for advanced operations not covered by the high-level workflow tools.")]
    public static async Task<object> ExecuteSolidWorksApi(
        ISolidWorksConnection connection,
        IApiDocumentationService apiDocs,
        [Description("Target interface name (e.g., 'ISketchManager', 'IFeatureManager', 'IModelDoc2')")] string interfaceName,
        [Description("Method name to call (e.g., 'CreateLine', 'FeatureExtrusion2')")] string methodName,
        [Description("Method parameters as JSON object with parameter names and values")] string? parametersJson = null,
        [Description("Include API documentation in response for context (default true)")] bool includeDocumentation = true)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            // Parse parameters
            Dictionary<string, object>? parameters = null;
            if (!string.IsNullOrEmpty(parametersJson))
            {
                try
                {
                    parameters = JsonSerializer.Deserialize<Dictionary<string, object>>(parametersJson);
                }
                catch (JsonException ex)
                {
                    return new { success = false, error = $"Invalid JSON parameters: {ex.Message}" };
                }
            }

            // Get documentation if requested
            ApiDocResult? documentation = null;
            if (includeDocumentation)
            {
                documentation = await apiDocs.GetMethodDocumentationAsync(interfaceName, methodName);
            }

            // Get the target interface instance
            var targetInterface = GetInterfaceInstance(connection, interfaceName);
            if (targetInterface == null)
            {
                return new
                {
                    success = false,
                    error = $"Could not get instance of interface: {interfaceName}",
                    documentation = documentation != null ? FormatDocumentation(documentation) : null
                };
            }

            // Find and invoke the method
            var result = InvokeMethod(targetInterface, methodName, parameters);

            return new
            {
                success = true,
                interfaceName = interfaceName,
                methodName = methodName,
                result = FormatResult(result),
                documentation = documentation != null ? FormatDocumentation(documentation) : null
            };
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message,
                stackTrace = ex.StackTrace
            };
        }
    }

    [McpServerTool]
    [Description("Gets or sets a property value on a SolidWorks interface")]
    public static async Task<object> AccessSolidWorksProperty(
        ISolidWorksConnection connection,
        IApiDocumentationService apiDocs,
        [Description("Target interface name (e.g., 'IModelDoc2', 'IFeature')")] string interfaceName,
        [Description("Property name to access")] string propertyName,
        [Description("Value to set (leave null to get current value)")] string? valueJson = null,
        [Description("Include API documentation in response (default true)")] bool includeDocumentation = true)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            // Get documentation
            ApiDocResult? documentation = null;
            if (includeDocumentation)
            {
                documentation = await apiDocs.GetMethodDocumentationAsync(interfaceName, propertyName);
            }

            // Get the target interface instance
            var targetInterface = GetInterfaceInstance(connection, interfaceName);
            if (targetInterface == null)
            {
                return new
                {
                    success = false,
                    error = $"Could not get instance of interface: {interfaceName}",
                    documentation = documentation != null ? FormatDocumentation(documentation) : null
                };
            }

            var type = targetInterface.GetType();

            // Set or get property using COM-compatible approach
            if (!string.IsNullOrEmpty(valueJson))
            {
                // Set property
                try
                {
                    var value = JsonSerializer.Deserialize<object>(valueJson);

                    // Try standard reflection first
                    var property = type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                    if (property != null && property.CanWrite)
                    {
                        property.SetValue(targetInterface, value);
                    }
                    else
                    {
                        // Fallback to InvokeMember for COM objects
                        type.InvokeMember(
                            propertyName,
                            BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                            null,
                            targetInterface,
                            new object[] { value! }
                        );
                    }

                    return new
                    {
                        success = true,
                        operation = "set",
                        interfaceName = interfaceName,
                        propertyName = propertyName,
                        value = value,
                        documentation = documentation != null ? FormatDocumentation(documentation) : null
                    };
                }
                catch (Exception ex)
                {
                    return new
                    {
                        success = false,
                        error = $"Failed to set property '{propertyName}': {ex.Message}"
                    };
                }
            }
            else
            {
                // Get property
                try
                {
                    object? value = null;

                    // Try standard reflection first
                    var property = type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                    if (property != null && property.CanRead)
                    {
                        value = property.GetValue(targetInterface);
                    }
                    else
                    {
                        // Fallback to InvokeMember for COM objects
                        value = type.InvokeMember(
                            propertyName,
                            BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                            null,
                            targetInterface,
                            null
                        );
                    }

                    return new
                    {
                        success = true,
                        operation = "get",
                        interfaceName = interfaceName,
                        propertyName = propertyName,
                        value = FormatResult(value),
                        documentation = documentation != null ? FormatDocumentation(documentation) : null
                    };
                }
                catch (Exception ex)
                {
                    return new
                    {
                        success = false,
                        error = $"Failed to get property '{propertyName}': {ex.Message}"
                    };
                }
            }
        }
        catch (Exception ex)
        {
            return new
            {
                success = false,
                error = ex.Message,
                stackTrace = ex.StackTrace
            };
        }
    }

    // Helper method to get interface instances
    private static object? GetInterfaceInstance(ISolidWorksConnection connection, string interfaceName)
    {
        var app = connection.Application;
        var doc = connection.ActiveDocument;

        // Common interface mappings
        return interfaceName.ToLowerInvariant() switch
        {
            "isldworks" or "sldworks" => app,
            "imodeldoc2" or "modeldoc2" => doc,
            "imodeldocextension" or "modeldocextension" => doc?.Extension,
            "isketchmanager" or "sketchmanager" => doc?.SketchManager,
            "ifeaturemanager" or "featuremanager" => doc?.FeatureManager,
            "iselectionmgr" or "selectionmgr" => doc?.SelectionManager,
            "ipartdoc" or "partdoc" => doc as IPartDoc,
            "iassemblydoc" or "assemblydoc" => doc as IAssemblyDoc,
            "idrawingdoc" or "drawingdoc" => doc as IDrawingDoc,
            _ => null
        };
    }

    // Helper method to invoke a method with parameters (COM-compatible hybrid approach)
    private static object? InvokeMethod(object targetInterface, string methodName, Dictionary<string, object>? parameters)
    {
        Type type = targetInterface.GetType();
        object?[] paramArray = parameters?.Values.ToArray() ?? Array.Empty<object>();

        // Approach 1: Try standard reflection first (works for non-COM objects)
        try
        {
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (methods.Length > 0)
            {
                // Try to find matching method by parameter count
                foreach (var method in methods)
                {
                    var methodParams = method.GetParameters();
                    var paramCount = parameters?.Count ?? 0;

                    // Try to match parameter count (accounting for optional parameters)
                    var requiredParams = methodParams.Count(p => !p.IsOptional);
                    if (paramCount >= requiredParams && paramCount <= methodParams.Length)
                    {
                        // Build parameter array
                        var invokeParams = new object?[methodParams.Length];
                        for (int i = 0; i < methodParams.Length; i++)
                        {
                            var param = methodParams[i];
                            if (parameters != null && parameters.TryGetValue(param.Name!, out var value))
                            {
                                invokeParams[i] = ConvertParameter(value, param.ParameterType);
                            }
                            else if (param.IsOptional)
                            {
                                invokeParams[i] = param.DefaultValue;
                            }
                            else
                            {
                                throw new InvalidOperationException($"Missing required parameter: {param.Name}");
                            }
                        }

                        return method.Invoke(targetInterface, invokeParams);
                    }
                }
            }
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            // Continue to COM approaches if standard reflection fails
        }

        // Approach 2: Try Type.InvokeMember (works better for COM objects)
        try
        {
            return type.InvokeMember(
                methodName,
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                targetInterface,
                paramArray
            );
        }
        catch
        {
            // Continue to next approach
        }

        // Approach 3: Try as property getter (for parameterless methods that might be properties)
        if (paramArray.Length == 0)
        {
            try
            {
                return type.InvokeMember(
                    methodName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    targetInterface,
                    null
                );
            }
            catch
            {
                // All approaches failed
            }
        }

        throw new InvalidOperationException(
            $"Method '{methodName}' could not be invoked on interface {type.Name}. " +
            "This may be a COM interop issue or the method does not exist. " +
            $"Attempted with {parameters?.Count ?? 0} parameter(s)."
        );
    }

    // Helper to convert parameter values to correct types
    private static object? ConvertParameter(object value, Type targetType)
    {
        if (value == null) return null;

        // Handle JSON element conversion
        if (value is JsonElement jsonElement)
        {
            return jsonElement.ValueKind switch
            {
                JsonValueKind.String => jsonElement.GetString(),
                JsonValueKind.Number => targetType == typeof(double) ? jsonElement.GetDouble() :
                                       targetType == typeof(int) ? jsonElement.GetInt32() :
                                       targetType == typeof(float) ? jsonElement.GetSingle() :
                                       jsonElement.GetDecimal(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => value
            };
        }

        return Convert.ChangeType(value, targetType);
    }

    // Helper to format results for JSON serialization
    private static object? FormatResult(object? result)
    {
        if (result == null) return null;

        var type = result.GetType();

        // Handle COM objects - extract basic info
        if (Marshal.IsComObject(result))
        {
            return new
            {
                type = type.Name,
                value = "COM Object (use GetModelInfo or other tools to inspect)"
            };
        }

        // Handle basic types
        if (type.IsPrimitive || type == typeof(string) || type == typeof(decimal))
        {
            return result;
        }

        // Handle arrays
        if (type.IsArray)
        {
            var array = (Array)result;
            return new
            {
                type = "Array",
                length = array.Length,
                values = array.Cast<object>().Take(10).Select(FormatResult).ToArray()
            };
        }

        // Default: return type info
        return new
        {
            type = type.Name,
            value = result.ToString()
        };
    }

    // Helper to format documentation
    private static object FormatDocumentation(ApiDocResult doc)
    {
        return new
        {
            interfaceName = doc.InterfaceName,
            methodName = doc.MethodName,
            description = doc.Description,
            syntax = doc.Syntax,
            remarks = doc.Remarks
        };
    }
}
