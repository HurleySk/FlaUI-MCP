using System.Text.Json;

namespace PlaywrightWindows.Mcp;

/// <summary>
/// Registry for MCP tools - maps tool names to handlers
/// </summary>
public class ToolRegistry
{
    private readonly Dictionary<string, ITool> _tools = new();

    public void RegisterTool(ITool tool)
    {
        _tools[tool.Name] = tool;
    }

    public List<McpTool> GetToolDefinitions()
    {
        return _tools.Values.Select(t => t.GetDefinition()).ToList();
    }

    private const int DefaultTimeoutMs = 30000;

    public async Task<McpToolResult> ExecuteToolAsync(string name, JsonElement? arguments)
    {
        if (!_tools.TryGetValue(name, out var tool))
        {
            return new McpToolResult
            {
                Content = new List<McpContent>
                {
                    new() { Type = "text", Text = $"Unknown tool: {name}" }
                },
                IsError = true
            };
        }

        try
        {
            var toolTask = tool.ExecuteAsync(arguments);
            var timeoutTask = Task.Delay(DefaultTimeoutMs);

            if (await Task.WhenAny(toolTask, timeoutTask) == timeoutTask)
            {
                // Observe the abandoned task's exception to avoid UnobservedTaskException
                _ = toolTask.ContinueWith(
                    t => { _ = t.Exception; },
                    TaskContinuationOptions.OnlyOnFaulted);

                return new McpToolResult
                {
                    Content = new List<McpContent>
                    {
                        new() { Type = "text", Text = $"Tool '{name}' timed out after {DefaultTimeoutMs}ms. " +
                            "A modal dialog (e.g., file Open/Save) may be blocking the UI. " +
                            "Try using windows_file_dialog to dismiss the dialog first." }
                    },
                    IsError = true
                };
            }

            return await toolTask;
        }
        catch (Exception ex)
        {
            return new McpToolResult
            {
                Content = new List<McpContent>
                {
                    new() { Type = "text", Text = $"Error: {ex.Message}" }
                },
                IsError = true
            };
        }
    }
}

/// <summary>
/// Interface for MCP tools
/// </summary>
public interface ITool
{
    string Name { get; }
    McpTool GetDefinition();
    Task<McpToolResult> ExecuteAsync(JsonElement? arguments);
}

/// <summary>
/// Base class for tools with common utilities
/// </summary>
public abstract class ToolBase : ITool
{
    public abstract string Name { get; }
    public abstract string Description { get; }
    public abstract object InputSchema { get; }

    public McpTool GetDefinition() => new()
    {
        Name = Name,
        Description = Description,
        InputSchema = InputSchema
    };

    public abstract Task<McpToolResult> ExecuteAsync(JsonElement? arguments);

    protected static McpToolResult TextResult(string text) => new()
    {
        Content = new List<McpContent>
        {
            new() { Type = "text", Text = text }
        }
    };

    protected static McpToolResult ErrorResult(string message) => new()
    {
        Content = new List<McpContent>
        {
            new() { Type = "text", Text = message }
        },
        IsError = true
    };

    protected static McpToolResult ImageResult(byte[] imageData, string mimeType = "image/png") => new()
    {
        Content = new List<McpContent>
        {
            new() 
            { 
                Type = "image", 
                Data = Convert.ToBase64String(imageData),
                MimeType = mimeType
            }
        }
    };

    protected T? GetArgument<T>(JsonElement? arguments, string name)
    {
        if (arguments == null) return default;
        if (!arguments.Value.TryGetProperty(name, out var prop)) return default;
        return JsonSerializer.Deserialize<T>(prop.GetRawText(), McpProtocol.JsonOptions);
    }

    protected string? GetStringArgument(JsonElement? arguments, string name)
    {
        if (arguments == null) return null;
        if (!arguments.Value.TryGetProperty(name, out var prop)) return null;
        return prop.GetString();
    }

    protected bool GetBoolArgument(JsonElement? arguments, string name, bool defaultValue = false)
    {
        if (arguments == null) return defaultValue;
        if (!arguments.Value.TryGetProperty(name, out var prop)) return defaultValue;
        return prop.GetBoolean();
    }
}
