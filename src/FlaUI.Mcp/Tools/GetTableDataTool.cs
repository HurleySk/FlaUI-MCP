using System.Text;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using PlaywrightWindows.Mcp.Core;

namespace PlaywrightWindows.Mcp.Tools;

/// <summary>
/// Extract structured data from DataGridView/Table controls
/// </summary>
public class GetTableDataTool : ToolBase
{
    private readonly ElementRegistry _elementRegistry;

    public GetTableDataTool(ElementRegistry elementRegistry)
    {
        _elementRegistry = elementRegistry;
    }

    public override string Name => "windows_get_table_data";

    public override string Description =>
        "Extract structured data from a DataGridView, Table, or Grid control. " +
        "Returns headers and rows as formatted text. Much cleaner than parsing the snapshot tree.";

    public override object InputSchema => new
    {
        type = "object",
        properties = new
        {
            @ref = new
            {
                type = "string",
                description = "Element ref of the grid/table control (from windows_snapshot or windows_find_elements)"
            },
            maxRows = new
            {
                type = "integer",
                description = "Maximum number of rows to return (default: 50)"
            }
        },
        required = new[] { "ref" }
    };

    public override Task<McpToolResult> ExecuteAsync(JsonElement? arguments)
    {
        var refId = GetStringArgument(arguments, "ref");
        if (string.IsNullOrEmpty(refId))
            return Task.FromResult(ErrorResult("Missing required argument: ref"));

        var maxRows = GetArgument<int?>(arguments, "maxRows") ?? 50;

        var element = _elementRegistry.GetElement(refId);
        if (element == null)
            return Task.FromResult(ErrorResult($"Element not found: {refId}. Run windows_snapshot to refresh element refs."));

        try
        {
            // Try Grid pattern first (most reliable for DataGridView)
            if (element.Patterns.Grid.IsSupported)
            {
                return Task.FromResult(ExtractViaGridPattern(element, maxRows));
            }

            // Try Table pattern
            if (element.Patterns.Table.IsSupported)
            {
                return Task.FromResult(ExtractViaTablePattern(element, maxRows));
            }

            // Fallback: walk the tree structure
            return Task.FromResult(ExtractViaTreeWalk(element, maxRows));
        }
        catch (Exception ex)
        {
            return Task.FromResult(ErrorResult($"Failed to extract table data: {ex.Message}"));
        }
    }

    private static McpToolResult ExtractViaGridPattern(AutomationElement element, int maxRows)
    {
        var grid = element.Patterns.Grid.Pattern;
        var rowCount = grid.RowCount.ValueOrDefault;
        var colCount = grid.ColumnCount.ValueOrDefault;

        if (rowCount == 0 || colCount == 0)
            return TextResult("Table is empty (0 rows or 0 columns).");

        // Try to get headers from Table pattern or first Header children
        var headers = GetHeaders(element, colCount);

        var sb = new StringBuilder();
        sb.AppendLine($"Table: {rowCount} rows x {colCount} columns");
        sb.AppendLine();

        // Header row
        if (headers.Count > 0)
        {
            sb.AppendLine("| " + string.Join(" | ", headers) + " |");
            sb.AppendLine("| " + string.Join(" | ", headers.Select(h => new string('-', Math.Max(h.Length, 3)))) + " |");
        }

        // Data rows
        var displayRows = Math.Min(rowCount, maxRows);
        for (int row = 0; row < displayRows; row++)
        {
            var cells = new List<string>();
            for (int col = 0; col < colCount; col++)
            {
                try
                {
                    var cell = grid.GetItem(row, col);
                    var value = GetCellValue(cell);
                    cells.Add(value);
                }
                catch
                {
                    cells.Add("");
                }
            }
            sb.AppendLine("| " + string.Join(" | ", cells) + " |");
        }

        if (rowCount > maxRows)
            sb.AppendLine($"... ({rowCount - maxRows} more rows)");

        return TextResult(sb.ToString());
    }

    private static McpToolResult ExtractViaTablePattern(AutomationElement element, int maxRows)
    {
        var table = element.Patterns.Table.Pattern;

        // Get column headers from Table pattern
        var headerElements = table.ColumnHeaders.ValueOrDefault;
        var headers = new List<string>();
        if (headerElements != null)
        {
            foreach (var h in headerElements)
            {
                headers.Add(h.Properties.Name.ValueOrDefault ?? "");
            }
        }

        // Use Grid pattern for row/col counts and cell access
        if (!element.Patterns.Grid.IsSupported)
        {
            // Table pattern without Grid — just return headers
            if (headers.Count > 0)
                return TextResult($"Table headers: {string.Join(", ", headers)} (Grid pattern not available for cell data)");
            return TextResult("Table pattern found but no Grid pattern for cell data.");
        }

        var grid = element.Patterns.Grid.Pattern;
        var rowCount = grid.RowCount.ValueOrDefault;
        var colCount = grid.ColumnCount.ValueOrDefault;

        if (rowCount == 0 || colCount == 0)
            return TextResult("Table is empty (0 rows or 0 columns).");

        var sb = new StringBuilder();
        sb.AppendLine($"Table: {rowCount} rows x {colCount} columns");
        sb.AppendLine();

        if (headers.Count > 0)
        {
            sb.AppendLine("| " + string.Join(" | ", headers) + " |");
            sb.AppendLine("| " + string.Join(" | ", headers.Select(h => new string('-', Math.Max(h.Length, 3)))) + " |");
        }

        var displayRows = Math.Min(rowCount, maxRows);
        for (int row = 0; row < displayRows; row++)
        {
            var cells = new List<string>();
            for (int col = 0; col < colCount; col++)
            {
                try
                {
                    var cell = grid.GetItem(row, col);
                    cells.Add(GetCellValue(cell));
                }
                catch
                {
                    cells.Add("");
                }
            }
            sb.AppendLine("| " + string.Join(" | ", cells) + " |");
        }

        if (rowCount > maxRows)
            sb.AppendLine($"... ({rowCount - maxRows} more rows)");

        return TextResult(sb.ToString());
    }

    private static McpToolResult ExtractViaTreeWalk(AutomationElement element, int maxRows)
    {
        // Walk children looking for Header and DataItem (row) elements
        var children = element.FindAllChildren();

        var headers = new List<string>();
        var rows = new List<List<string>>();

        foreach (var child in children)
        {
            var controlType = child.Properties.ControlType.ValueOrDefault;

            if (controlType == ControlType.Header)
            {
                var headerItems = child.FindAllChildren();
                foreach (var hi in headerItems)
                {
                    headers.Add(hi.Properties.Name.ValueOrDefault ?? "");
                }
            }
            else if (controlType == ControlType.DataItem && rows.Count < maxRows)
            {
                var cellElements = child.FindAllChildren();
                var cells = new List<string>();
                foreach (var cell in cellElements)
                {
                    cells.Add(GetCellValue(cell));
                }
                rows.Add(cells);
            }
        }

        if (rows.Count == 0 && headers.Count == 0)
            return TextResult("No table data found. The element may not be a table/grid control.");

        var sb = new StringBuilder();
        sb.AppendLine($"Table: {rows.Count} rows");
        sb.AppendLine();

        if (headers.Count > 0)
        {
            sb.AppendLine("| " + string.Join(" | ", headers) + " |");
            sb.AppendLine("| " + string.Join(" | ", headers.Select(h => new string('-', Math.Max(h.Length, 3)))) + " |");
        }

        foreach (var row in rows)
        {
            sb.AppendLine("| " + string.Join(" | ", row) + " |");
        }

        return TextResult(sb.ToString());
    }

    private static List<string> GetHeaders(AutomationElement gridElement, int colCount)
    {
        var headers = new List<string>();

        // Try Table pattern for headers
        if (gridElement.Patterns.Table.IsSupported)
        {
            var headerElements = gridElement.Patterns.Table.Pattern.ColumnHeaders.ValueOrDefault;
            if (headerElements != null)
            {
                foreach (var h in headerElements)
                {
                    headers.Add(h.Properties.Name.ValueOrDefault ?? "");
                }
                if (headers.Count > 0) return headers;
            }
        }

        // Fallback: find Header children
        try
        {
            var headerRow = gridElement.FindFirstChild(cf => cf.ByControlType(ControlType.Header));
            if (headerRow != null)
            {
                var headerItems = headerRow.FindAllChildren();
                foreach (var hi in headerItems)
                {
                    headers.Add(hi.Properties.Name.ValueOrDefault ?? "");
                }
            }
        }
        catch
        {
            // Generate placeholder headers
            for (int i = 0; i < colCount; i++)
                headers.Add($"Col{i + 1}");
        }

        return headers;
    }

    private static string GetCellValue(AutomationElement cell)
    {
        try
        {
            // Try Value pattern
            if (cell.Patterns.Value.IsSupported)
            {
                var value = cell.Patterns.Value.Pattern.Value.ValueOrDefault;
                if (!string.IsNullOrEmpty(value)) return value;
            }

            // Try Name property
            var name = cell.Properties.Name.ValueOrDefault;
            if (!string.IsNullOrEmpty(name)) return name;

            // Try first text child
            var textChild = cell.FindFirstChild(cf => cf.ByControlType(ControlType.Text));
            if (textChild != null)
            {
                var textName = textChild.Properties.Name.ValueOrDefault ?? "";
                if (!string.IsNullOrEmpty(textName)) return textName;
            }

            // Try LegacyIAccessible pattern (works for WinForms DataGridView)
            if (cell.Patterns.LegacyIAccessible.IsSupported)
            {
                var legacyValue = cell.Patterns.LegacyIAccessible.Pattern.Value.ValueOrDefault;
                if (!string.IsNullOrEmpty(legacyValue)) return legacyValue;
                var legacyName = cell.Patterns.LegacyIAccessible.Pattern.Name.ValueOrDefault;
                if (!string.IsNullOrEmpty(legacyName)) return legacyName;
            }

            return "";
        }
        catch
        {
            return "";
        }
    }
}
