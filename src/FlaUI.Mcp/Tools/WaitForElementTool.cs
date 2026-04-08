using System.Diagnostics;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using PlaywrightWindows.Mcp.Core;

namespace PlaywrightWindows.Mcp.Tools;

/// <summary>
/// Wait for an element to meet a condition (exists, enabled, visible, has text)
/// </summary>
public class WaitForElementTool : ToolBase
{
    private readonly SessionManager _sessionManager;
    private readonly ElementRegistry _elementRegistry;

    public WaitForElementTool(SessionManager sessionManager, ElementRegistry elementRegistry)
    {
        _sessionManager = sessionManager;
        _elementRegistry = elementRegistry;
    }

    public override string Name => "windows_wait_for_element";

    public override string Description =>
        "Wait for an element matching name or automationId to meet a condition. " +
        "Polls the accessibility tree until the condition is met or timeout. " +
        "Returns the element ref when found. Use instead of arbitrary waits after clicking buttons.";

    public override object InputSchema => new
    {
        type = "object",
        properties = new
        {
            handle = new
            {
                type = "string",
                description = "Window handle to search in"
            },
            name = new
            {
                type = "string",
                description = "Element Name property to match (substring match)"
            },
            automationId = new
            {
                type = "string",
                description = "Element AutomationId to match (exact match)"
            },
            condition = new
            {
                type = "string",
                @enum = new[] { "exists", "enabled", "visible", "has_text" },
                description = "Condition to wait for (default: exists)"
            },
            timeout = new
            {
                type = "integer",
                description = "Max milliseconds to wait (default: 10000)"
            }
        },
        required = new[] { "handle" }
    };

    public override async Task<McpToolResult> ExecuteAsync(JsonElement? arguments)
    {
        var handle = GetStringArgument(arguments, "handle");
        if (string.IsNullOrEmpty(handle))
            return ErrorResult("Missing required argument: handle");

        var name = GetStringArgument(arguments, "name");
        var automationId = GetStringArgument(arguments, "automationId");
        var condition = GetStringArgument(arguments, "condition") ?? "exists";
        var timeout = GetArgument<int?>(arguments, "timeout") ?? 10000;

        if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(automationId))
            return ErrorResult("At least one of 'name' or 'automationId' must be provided");

        var window = _sessionManager.GetWindow(handle);
        if (window == null)
            return ErrorResult($"Window not found: {handle}");

        try
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeout)
            {
                var match = FindMatchingElement(window, name, automationId);
                if (match != null && MeetsCondition(match, condition))
                {
                    var refId = _elementRegistry.Register(handle, match);
                    var elementName = match.Properties.Name.ValueOrDefault ?? automationId ?? "element";
                    var elapsed = sw.ElapsedMilliseconds;
                    return TextResult($"Found: {elementName} [ref={refId}] (condition '{condition}' met after {elapsed}ms)");
                }

                await Task.Delay(250);
            }

            return ErrorResult($"Timeout after {timeout}ms waiting for element (name={name ?? "any"}, automationId={automationId ?? "any"}) with condition '{condition}'");
        }
        catch (Exception ex)
        {
            return ErrorResult($"Failed waiting for element: {ex.Message}");
        }
    }

    private static AutomationElement? FindMatchingElement(AutomationElement root, string? name, string? automationId)
    {
        try
        {
            var descendants = root.FindAllDescendants();
            foreach (var el in descendants)
            {
                try
                {
                    if (!string.IsNullOrEmpty(automationId))
                    {
                        var elAutoId = el.Properties.AutomationId.ValueOrDefault;
                        if (!string.Equals(elAutoId, automationId, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }

                    if (!string.IsNullOrEmpty(name))
                    {
                        var elName = el.Properties.Name.ValueOrDefault ?? "";
                        if (!elName.Contains(name, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }

                    return el;
                }
                catch
                {
                    // Skip inaccessible elements
                }
            }
        }
        catch
        {
            // Tree walk failed
        }

        return null;
    }

    private static bool MeetsCondition(AutomationElement element, string condition)
    {
        try
        {
            return condition switch
            {
                "exists" => true,
                "enabled" => element.Properties.IsEnabled.ValueOrDefault,
                "visible" => !element.Properties.IsOffscreen.ValueOrDefault,
                "has_text" => HasText(element),
                _ => true
            };
        }
        catch
        {
            return false;
        }
    }

    private static bool HasText(AutomationElement element)
    {
        try
        {
            if (element.Patterns.Value.IsSupported)
            {
                var value = element.Patterns.Value.Pattern.Value.ValueOrDefault;
                if (!string.IsNullOrEmpty(value)) return true;
            }

            var name = element.Properties.Name.ValueOrDefault;
            return !string.IsNullOrEmpty(name);
        }
        catch
        {
            return false;
        }
    }
}
