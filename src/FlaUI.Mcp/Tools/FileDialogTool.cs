using System.Diagnostics;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using PlaywrightWindows.Mcp.Core;

namespace PlaywrightWindows.Mcp.Tools;

/// <summary>
/// Automate Windows file open/save dialogs
/// </summary>
public class FileDialogTool : ToolBase
{
    private readonly SessionManager _sessionManager;

    public FileDialogTool(SessionManager sessionManager)
    {
        _sessionManager = sessionManager;
    }

    public override string Name => "windows_file_dialog";

    public override string Description =>
        "Complete a Windows file Open/Save dialog by filling in the file path and clicking Open/Save. " +
        "Call this AFTER clicking a Browse button that opens a file dialog. " +
        "Waits for the dialog to appear, fills the filename field, and confirms.";

    public override object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Full file path to enter in the dialog (e.g., 'C:\\data\\input.csv')"
            },
            action = new
            {
                type = "string",
                @enum = new[] { "open", "save" },
                description = "Dialog type: 'open' or 'save' (default: open)"
            },
            timeout = new
            {
                type = "integer",
                description = "Max milliseconds to wait for dialog to appear (default: 5000)"
            }
        },
        required = new[] { "path" }
    };

    public override async Task<McpToolResult> ExecuteAsync(JsonElement? arguments)
    {
        var path = GetStringArgument(arguments, "path");
        if (string.IsNullOrEmpty(path))
            return ErrorResult("Missing required argument: path");

        var action = GetStringArgument(arguments, "action") ?? "open";
        var timeout = GetArgument<int?>(arguments, "timeout") ?? 5000;

        try
        {
            var dialog = await WaitForFileDialog(timeout);
            if (dialog == null)
                return ErrorResult($"No file dialog appeared within {timeout}ms. Make sure you clicked a Browse/Open button first.");

            // Brief delay to let the dialog fully initialize its UIA tree
            await Task.Delay(200);
            dialog.Focus();
            await Task.Delay(100);

            // Strategy 1: Find the filename edit control by AutomationId "1148" (standard Windows file dialog)
            // Retry a few times since modal dialog controls may not be immediately available
            AutomationElement? filenameEdit = null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                filenameEdit = FindFilenameEdit(dialog);
                if (filenameEdit != null) break;
                await Task.Delay(300);
            }
            if (filenameEdit != null)
            {
                // Clear and fill
                if (filenameEdit.Patterns.Value.IsSupported)
                {
                    filenameEdit.Patterns.Value.Pattern.SetValue(path);
                }
                else
                {
                    filenameEdit.Focus();
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    await Task.Delay(50);
                    Keyboard.Type(path);
                }

                await Task.Delay(100);

                // Find and click Open/Save button
                var confirmButton = FindConfirmButton(dialog, action);
                if (confirmButton != null)
                {
                    if (confirmButton.Patterns.Invoke.IsSupported)
                        confirmButton.Patterns.Invoke.Pattern.Invoke();
                    else
                        confirmButton.Click();

                    return TextResult($"File dialog completed: {path}");
                }

                // Fallback: press Enter
                Keyboard.Press(VirtualKeyShort.ENTER);
                return TextResult($"File dialog completed (Enter key): {path}");
            }

            // Strategy 2: Keyboard-based fallback
            dialog.Focus();
            await Task.Delay(100);

            // Type the path directly — most file dialogs accept keyboard input in the filename field
            Keyboard.Type(path);
            await Task.Delay(100);
            Keyboard.Press(VirtualKeyShort.ENTER);

            return TextResult($"File dialog completed (keyboard fallback): {path}");
        }
        catch (Exception ex)
        {
            return ErrorResult($"Failed to complete file dialog: {ex.Message}");
        }
    }

    private async Task<Window?> WaitForFileDialog(int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var desktop = _sessionManager.Automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf =>
                cf.ByControlType(ControlType.Window));

            foreach (var win in windows)
            {
                try
                {
                    var className = win.Properties.ClassName.ValueOrDefault;
                    var name = win.Properties.Name.ValueOrDefault ?? "";

                    // Standard Windows file dialog class
                    if (className == "#32770")
                        return win.AsWindow();

                    // Also match by common dialog titles
                    if (name.Contains("Open") || name.Contains("Save") || name.Contains("Browse"))
                    {
                        if (className?.Contains("Dialog") == true || className == "#32770")
                            return win.AsWindow();
                    }
                }
                catch
                {
                    // Skip inaccessible windows
                }
            }

            await Task.Delay(200);
        }

        return null;
    }

    private static AutomationElement? FindFilenameEdit(AutomationElement dialog)
    {
        try
        {
            // Standard file dialog: filename edit has AutomationId "1148"
            var edits = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Edit));
            foreach (var edit in edits)
            {
                var autoId = edit.Properties.AutomationId.ValueOrDefault;
                if (autoId == "1148" || autoId == "1001")
                    return edit;
            }

            // Fallback: find the combo box with AutomationId "1148" and get its edit child
            var combos = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.ComboBox));
            foreach (var combo in combos)
            {
                var autoId = combo.Properties.AutomationId.ValueOrDefault;
                if (autoId == "1148" || autoId == "1001")
                {
                    var innerEdit = combo.FindFirstChild(cf => cf.ByControlType(ControlType.Edit));
                    if (innerEdit != null) return innerEdit;
                    return combo;
                }
            }

            // Last resort: first visible edit
            return edits.FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private static AutomationElement? FindConfirmButton(AutomationElement dialog, string action)
    {
        try
        {
            var buttons = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
            var targetName = action == "save" ? "Save" : "Open";

            // Exact match first
            foreach (var btn in buttons)
            {
                var name = btn.Properties.Name.ValueOrDefault ?? "";
                if (name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    return btn;
            }

            // AutomationId "1" is the default OK/Open button
            foreach (var btn in buttons)
            {
                var autoId = btn.Properties.AutomationId.ValueOrDefault;
                if (autoId == "1")
                    return btn;
            }

            return null;
        }
        catch
        {
            return null;
        }
    }
}
