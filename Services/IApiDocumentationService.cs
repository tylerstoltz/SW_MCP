namespace SolidWorksMCP.Services;

/// <summary>
/// Service for searching and retrieving SolidWorks API documentation
/// </summary>
public interface IApiDocumentationService
{
    /// <summary>
    /// Searches the SolidWorks API documentation for the given query
    /// </summary>
    /// <param name="query">Search query (interface name, method name, keywords, etc.)</param>
    /// <param name="maxResults">Maximum number of results to return</param>
    /// <returns>List of matching documentation entries with context</returns>
    Task<List<ApiDocResult>> SearchAsync(string query, int maxResults = 10);

    /// <summary>
    /// Gets documentation for a specific interface
    /// </summary>
    Task<ApiDocResult?> GetInterfaceDocumentationAsync(string interfaceName);

    /// <summary>
    /// Gets documentation for a specific method on an interface
    /// </summary>
    Task<ApiDocResult?> GetMethodDocumentationAsync(string interfaceName, string methodName);

    /// <summary>
    /// Gets code examples related to a query
    /// </summary>
    Task<List<CodeExample>> GetCodeExamplesAsync(string query, int maxResults = 5);
}

/// <summary>
/// Represents a documentation search result
/// </summary>
public class ApiDocResult
{
    public string InterfaceName { get; set; } = string.Empty;
    public string MethodName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Syntax { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public double RelevanceScore { get; set; }
}

/// <summary>
/// Represents a code example from the documentation
/// </summary>
public class CodeExample
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Code { get; set; } = string.Empty;
    public string Language { get; set; } = "csharp";
    public string FilePath { get; set; } = string.Empty;
}
