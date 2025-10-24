using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SolidWorksMCP.Services;

namespace SolidWorksMCP;

class Program
{
    static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);

        // CRITICAL: Configure logging to write to stderr only (stdout is reserved for JSON-RPC)
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(options =>
        {
            options.LogToStandardErrorThreshold = LogLevel.Trace;  // All logs to stderr
        });

        // Register SolidWorks connection as singleton
        builder.Services.AddSingleton<ISolidWorksConnection, SolidWorksConnection>();

        // Register API documentation search service
        builder.Services.AddSingleton<IApiDocumentationService, ApiDocumentationService>();

        // Configure MCP server with all capabilities
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly()        // Auto-discover all tools
            .WithPromptsFromAssembly()      // Auto-discover all prompts
            .WithResourcesFromAssembly();   // Auto-discover all resources

        var host = builder.Build();

        // Initialize SolidWorks connection on startup
        var swConnection = host.Services.GetRequiredService<ISolidWorksConnection>();
        await swConnection.InitializeAsync();

        await host.RunAsync();
    }
}
