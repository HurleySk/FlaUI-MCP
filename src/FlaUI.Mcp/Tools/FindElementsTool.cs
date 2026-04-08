using System.Text;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using PlaywrightWindows.Mcp.Core;

namespace PlaywrightWindows.Mcp.Tools;

/// <summary>
/// Search for elements matching criteria without a full snapshot
/// </summary>
public class FindElementsTool : ToolBase
{
    private readonly SessionManager _sessionManager;
    private readonly ElementRegistry _elementRegistry;

    public FindElementsTool(SessionManager sessionManager, ElementRegistry elementRegistry)
    {
        _sessionManager = sessionManager;
        _elementRegistry = elementRegistry;
    }

    public override string Name => "windows_find_elements";

    public override string Description =>
        "Search for elements matching name, automationId, or control type. " +
        "More efficient than windows_snapshot when you know what you're looking for. " +
        "Returns matching elements with their refs.";

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
                description = "Element Name to match (substring, case-insensitive)"
            },
            automationId = new
            {
                type = "string",
                description = "Element AutomationId to match (exact, case-insensitive)"
            },
            controlType = new
            {
                type = "string",
                description = "UI Automation control type (e.g., 'Button', 'Edit', 'DataGrid', 'ComboBox', 'CheckBox', 'List', 'Table', 'Text')"
            },
            nameContains = new
            {
                type = "string",
                description = "Substring to match in element Name (case-insensitive)"
            },
            maxResults = new
            {
                type = "integer",
                description = "Maximum number of results to return (default: 25)"
            }
        },
        required = new[] { "handle" }
    };

    public override Task<McpToolResult> ExecuteAsync(JsonElement? arguments)
    {
        var handle = GetStringArgument(arguments, "handle");
        if (string.IsNullOrEmpty(handle))
            return Task.FromResult(ErrorResult("Missing required argument: handle"));

        var name = GetStringArgument(arguments, "name");
        var automationId = GetStringArgument(arguments, "automationId");
        var controlType = GetStringArgument(arguments, "controlType");
        var nameContains = GetStringArgument(arguments, "nameContains");
        var maxResults = GetArgument<int?>(arguments, "maxResults") ?? 25;

        if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(automationId)
            && string.IsNullOrEmpty(controlType) && string.IsNullOrEmpty(nameContains))
        {
            return Task.FromResult(ErrorResult("At least one search criterion (name, automationId, controlType, nameContains) is required"));
        }

        var window = _sessionManager.GetWindow(handle);
        if (window == null)
            return Task.FromResult(ErrorResult($"Window not found: {handle}"));

        try
        {
            var parsedControlType = ParseControlType(controlType);
            var matches = new List<(AutomationElement element, string refId)>();

            var descendants = window.FindAllDescendants();
            foreach (var el in descendants)
            {
                if (matches.Count >= maxResults) break;

                try
                {
                    if (!MatchesFilter(el, name, automationId, parsedControlType, nameContains))
                        continue;

                    var refId = _elementRegistry.Register(handle, el);
                    matches.Add((el, refId));
                }
                catch
                {
                    // Skip inaccessible elements
                }
            }

            if (matches.Count == 0)
                return Task.FromResult(TextResult("No elements found matching the criteria."));

            var sb = new StringBuilder();
            sb.AppendLine($"Found {matches.Count} element(s):");
            foreach (var (el, refId) in matches)
            {
                var elName = el.Properties.Name.ValueOrDefault ?? "";
                var elAutoId = el.Properties.AutomationId.ValueOrDefault ?? "";
                var elType = el.Properties.ControlType.ValueOrDefault.ToString();
                var enabled = el.Properties.IsEnabled.ValueOrDefault ? "" : " [disabled]";

                sb.AppendLine($"  - [ref={refId}] {elType} name=\"{elName}\" automationId=\"{elAutoId}\"{enabled}");
            }

            return Task.FromResult(TextResult(sb.ToString()));
        }
        catch (Exception ex)
        {
            return Task.FromResult(ErrorResult($"Failed to find elements: {ex.Message}"));
        }
    }

    private static bool MatchesFilter(AutomationElement element, string? name, string? automationId,
        ControlType? controlType, string? nameContains)
    {
        if (controlType.HasValue)
        {
            var elType = element.Properties.ControlType.ValueOrDefault;
            if (elType != controlType.Value) return false;
        }

        if (!string.IsNullOrEmpty(automationId))
        {
            var elAutoId = element.Properties.AutomationId.ValueOrDefault ?? "";
            if (!elAutoId.Equals(automationId, StringComparison.OrdinalIgnoreCase)) return false;
        }

        if (!string.IsNullOrEmpty(name))
        {
            var elName = element.Properties.Name.ValueOrDefault ?? "";
            if (!elName.Equals(name, StringComparison.OrdinalIgnoreCase)) return false;
        }

        if (!string.IsNullOrEmpty(nameContains))
        {
            var elName = element.Properties.Name.ValueOrDefault ?? "";
            if (!elName.Contains(nameContains, StringComparison.OrdinalIgnoreCase)) return false;
        }

        return true;
    }

    private static ControlType? ParseControlType(string? controlTypeName)
    {
        if (string.IsNullOrEmpty(controlTypeName)) return null;

        return controlTypeName.ToLowerInvariant() switch
        {
            "button" => ControlType.Button,
            "edit" or "textbox" => ControlType.Edit,
            "text" or "label" => ControlType.Text,
            "checkbox" => ControlType.CheckBox,
            "radiobutton" or "radio" => ControlType.RadioButton,
            "combobox" or "dropdown" => ControlType.ComboBox,
            "list" => ControlType.List,
            "listitem" => ControlType.ListItem,
            "menu" => ControlType.Menu,
            "menuitem" => ControlType.MenuItem,
            "tree" => ControlType.Tree,
            "treeitem" => ControlType.TreeItem,
            "tab" or "tabcontrol" => ControlType.Tab,
            "tabitem" or "tabpage" => ControlType.TabItem,
            "table" => ControlType.Table,
            "datagrid" or "grid" => ControlType.DataGrid,
            "dataitem" or "row" => ControlType.DataItem,
            "header" => ControlType.Header,
            "headeritem" or "columnheader" => ControlType.HeaderItem,
            "image" => ControlType.Image,
            "slider" => ControlType.Slider,
            "spinner" or "spinbutton" => ControlType.Spinner,
            "progressbar" => ControlType.ProgressBar,
            "hyperlink" or "link" => ControlType.Hyperlink,
            "toolbar" => ControlType.ToolBar,
            "statusbar" => ControlType.StatusBar,
            "window" => ControlType.Window,
            "pane" or "group" => ControlType.Pane,
            "document" => ControlType.Document,
            "custom" => ControlType.Custom,
            _ => null
        };
    }
}
