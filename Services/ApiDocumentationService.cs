using System.Text.RegularExpressions;
using System.Text.Json;

namespace SolidWorksMCP.Services;

/// <summary>
/// Service for searching and retrieving SolidWorks API documentation
/// </summary>
public class ApiDocumentationService : IApiDocumentationService
{
    private readonly string _apiDocsPath;
    private readonly string _htmlDocsPath;
    private List<string>? _htmlFiles;

    public ApiDocumentationService()
    {
        _apiDocsPath = @"Path\\to\\solidworks\\Solidworks-SDK\\docs\\api\\";
        _htmlDocsPath = Path.Combine(_apiDocsPath, "html");
    }

    public async Task<List<ApiDocResult>> SearchAsync(string query, int maxResults = 10)
    {
        if (_htmlFiles == null)
        {
            await IndexDocumentationAsync();
        }

        var results = new List<ApiDocResult>();
        var queryLower = query.ToLowerInvariant();
        var queryTerms = queryLower.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        foreach (var file in _htmlFiles!)
        {
            try
            {
                var content = await File.ReadAllTextAsync(file);
                var score = CalculateRelevanceScore(content, file, queryTerms);

                if (score > 0)
                {
                    var result = ParseHtmlDocumentation(content, file);
                    result.RelevanceScore = score;
                    results.Add(result);
                }

                if (results.Count >= maxResults * 3) // Pre-filter to avoid processing all files
                    break;
            }
            catch (Exception ex)
            {
                // Skip files that can't be read
                Console.Error.WriteLine($"Error reading {file}: {ex.Message}");
            }
        }

        return results
            .OrderByDescending(r => r.RelevanceScore)
            .Take(maxResults)
            .ToList();
    }

    public async Task<ApiDocResult?> GetInterfaceDocumentationAsync(string interfaceName)
    {
        var results = await SearchAsync(interfaceName, 5);
        return results.FirstOrDefault(r =>
            r.InterfaceName.Equals(interfaceName, StringComparison.OrdinalIgnoreCase));
    }

    public async Task<ApiDocResult?> GetMethodDocumentationAsync(string interfaceName, string methodName)
    {
        var results = await SearchAsync($"{interfaceName} {methodName}", 5);
        return results.FirstOrDefault(r =>
            r.InterfaceName.Equals(interfaceName, StringComparison.OrdinalIgnoreCase) &&
            r.MethodName.Equals(methodName, StringComparison.OrdinalIgnoreCase));
    }

    public async Task<List<CodeExample>> GetCodeExamplesAsync(string query, int maxResults = 5)
    {
        if (_htmlFiles == null)
        {
            await IndexDocumentationAsync();
        }

        var examples = new List<CodeExample>();
        var queryLower = query.ToLowerInvariant();

        // Look for files with "_Example_" in the name (common pattern in SW docs)
        var exampleFiles = _htmlFiles!
            .Where(f => f.Contains("_Example_", StringComparison.OrdinalIgnoreCase) ||
                        f.Contains("_CSharp", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var file in exampleFiles)
        {
            try
            {
                var content = await File.ReadAllTextAsync(file);
                if (content.ToLowerInvariant().Contains(queryLower))
                {
                    var example = ParseCodeExample(content, file);
                    if (!string.IsNullOrEmpty(example.Code))
                    {
                        examples.Add(example);
                    }

                    if (examples.Count >= maxResults)
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error reading example {file}: {ex.Message}");
            }
        }

        return examples;
    }

    private async Task IndexDocumentationAsync()
    {
        await Task.Run(() =>
        {
            if (Directory.Exists(_htmlDocsPath))
            {
                _htmlFiles = Directory.GetFiles(_htmlDocsPath, "*.htm", SearchOption.AllDirectories).ToList();
                Console.Error.WriteLine($"Indexed {_htmlFiles.Count} HTML documentation files");
            }
            else
            {
                _htmlFiles = new List<string>();
                Console.Error.WriteLine($"Warning: Documentation path not found: {_htmlDocsPath}");
            }
        });
    }

    private double CalculateRelevanceScore(string content, string filePath, string[] queryTerms)
    {
        var contentLower = content.ToLowerInvariant();
        var fileNameLower = Path.GetFileName(filePath).ToLowerInvariant();
        double score = 0;

        foreach (var term in queryTerms)
        {
            // File name matches are highest priority
            if (fileNameLower.Contains(term))
                score += 10;

            // Count occurrences in content
            var count = Regex.Matches(contentLower, Regex.Escape(term)).Count;
            score += count * 0.5;
        }

        // Boost interface documentation files
        if (fileNameLower.StartsWith("i") && !fileNameLower.Contains("example"))
            score *= 1.5;

        return score;
    }

    private ApiDocResult ParseHtmlDocumentation(string html, string filePath)
    {
        var result = new ApiDocResult { FilePath = filePath };

        // Extract interface/class name from file name or content
        var fileName = Path.GetFileNameWithoutExtension(filePath);
        result.InterfaceName = ExtractInterfaceName(fileName, html);
        result.MethodName = ExtractMethodName(fileName, html);

        // Extract description - look for common patterns in SW API docs
        result.Description = ExtractSection(html, "Description", 500);
        if (string.IsNullOrEmpty(result.Description))
        {
            result.Description = ExtractFirstParagraph(html, 500);
        }

        // Extract syntax
        result.Syntax = ExtractSection(html, "Syntax", 1000);

        // Extract remarks
        result.Remarks = ExtractSection(html, "Remarks", 1000);

        return result;
    }

    private string ExtractInterfaceName(string fileName, string html)
    {
        // Try to extract from common patterns
        if (fileName.StartsWith("I") && fileName.Contains("_"))
        {
            return fileName.Split('_')[0];
        }

        // Try to extract from HTML content
        var match = Regex.Match(html, @"interface\s+([A-Z][a-zA-Z0-9]+)", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        return fileName;
    }

    private string ExtractMethodName(string fileName, string html)
    {
        // Extract method name from file name pattern like "IInterface_MethodName"
        var parts = fileName.Split('_');
        if (parts.Length >= 2 && !parts[1].Contains("Example"))
        {
            return parts[1];
        }

        return string.Empty;
    }

    private string ExtractSection(string html, string sectionName, int maxLength)
    {
        // Look for common heading patterns in SolidWorks API docs
        var patterns = new[]
        {
            $@"<h[234]>\s*{sectionName}\s*</h[234]>\s*<[^>]+>([^<]+)",
            $@"<strong>\s*{sectionName}\s*[:\s]*</strong>([^<]+)",
            $@"{sectionName}:\s*</[^>]+>\s*<[^>]+>([^<]+)"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (match.Success)
            {
                var text = StripHtmlTags(match.Groups[1].Value).Trim();
                return text.Length > maxLength ? text.Substring(0, maxLength) + "..." : text;
            }
        }

        return string.Empty;
    }

    private string ExtractFirstParagraph(string html, int maxLength)
    {
        var match = Regex.Match(html, @"<p[^>]*>(.+?)</p>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (match.Success)
        {
            var text = StripHtmlTags(match.Groups[1].Value).Trim();
            return text.Length > maxLength ? text.Substring(0, maxLength) + "..." : text;
        }

        return string.Empty;
    }

    private CodeExample ParseCodeExample(string html, string filePath)
    {
        var example = new CodeExample
        {
            FilePath = filePath,
            Title = Path.GetFileNameWithoutExtension(filePath).Replace("_", " "),
            Language = filePath.Contains("CSharp", StringComparison.OrdinalIgnoreCase) ? "csharp" : "vb"
        };

        // Extract code blocks - look for <pre> or <code> tags
        var codeMatch = Regex.Match(html, @"<pre[^>]*>(.*?)</pre>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (!codeMatch.Success)
        {
            codeMatch = Regex.Match(html, @"<code[^>]*>(.*?)</code>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        }

        if (codeMatch.Success)
        {
            example.Code = StripHtmlTags(codeMatch.Groups[1].Value).Trim();
        }

        // Extract description
        example.Description = ExtractFirstParagraph(html, 300);

        return example;
    }

    private string StripHtmlTags(string html)
    {
        var text = Regex.Replace(html, @"<[^>]+>", " ");
        text = Regex.Replace(text, @"\s+", " ");
        return System.Net.WebUtility.HtmlDecode(text);
    }
}
