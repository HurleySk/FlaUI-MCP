using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
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
            var (dialog, hwnd) = await WaitForFileDialog(timeout);
            if (dialog == null)
                return ErrorResult($"No file dialog appeared within {timeout}ms. Make sure you clicked a Browse/Open button first.");

            // Brief delay to let the dialog fully initialize its UIA tree
            await Task.Delay(200);

            // Use SetForegroundWindow via native HWND — more reliable than UIA Focus()
            // when the parent app's UI thread is blocked by a modal dialog
            if (hwnd != IntPtr.Zero)
                NativeMethods.SetForegroundWindow(hwnd);
            else
                dialog.Focus();
            await Task.Delay(100);

            // Strategy 1: Find the filename edit control by AutomationId "1148" (standard Windows file dialog)
            // Retry a few times since modal dialog controls may not be immediately available.
            // Wrap in Task.Run with timeout to protect against UIA hangs.
            AutomationElement? filenameEdit = null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                var findTask = Task.Run(() => FindFilenameEdit(dialog));
                if (await Task.WhenAny(findTask, Task.Delay(3000)) == findTask)
                    filenameEdit = await findTask;
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

                // Find and click Open/Save button (with timeout protection)
                AutomationElement? confirmButton = null;
                var btnTask = Task.Run(() => FindConfirmButton(dialog, action));
                if (await Task.WhenAny(btnTask, Task.Delay(3000)) == btnTask)
                    confirmButton = await btnTask;

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

            // Strategy 2: Keyboard-based fallback — use native HWND focus
            if (hwnd != IntPtr.Zero)
                NativeMethods.SetForegroundWindow(hwnd);
            else
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

    private async Task<(Window? dialog, IntPtr hwnd)> WaitForFileDialog(int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            // Strategy A: Try UIA tree traversal with a short inner timeout.
            // This is the fast path for non-WinForms dialogs where UIA works fine.
            var uiaTask = Task.Run(() =>
            {
                try
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

                            if (className == "#32770")
                                return win.AsWindow();

                            if (name.Contains("Open") || name.Contains("Save") || name.Contains("Browse"))
                            {
                                if (className?.Contains("Dialog") == true || className == "#32770")
                                    return win.AsWindow();
                            }
                        }
                        catch { /* Skip inaccessible windows */ }
                    }
                }
                catch { /* UIA traversal failed entirely */ }
                return null;
            });

            if (await Task.WhenAny(uiaTask, Task.Delay(2000)) == uiaTask)
            {
                var result = await uiaTask;
                if (result != null)
                    return (result, IntPtr.Zero);
            }

            // Strategy B: Win32 fallback — EnumWindows bypasses UIA entirely.
            // WinForms modal dialogs block UIA desktop traversal but are still
            // discoverable via Win32 and can be wrapped with Automation.FromHandle().
            var dialogHwnd = FindDialogHwndViaWin32();
            if (dialogHwnd != IntPtr.Zero)
            {
                try
                {
                    var element = _sessionManager.Automation.FromHandle(dialogHwnd);
                    if (element != null)
                        return (element.AsWindow(), dialogHwnd);
                }
                catch { /* FromHandle can fail if the dialog is closing */ }
            }

            await Task.Delay(200);
        }

        return (null, IntPtr.Zero);
    }

    /// <summary>
    /// Find a file dialog window handle using Win32 EnumWindows.
    /// This works even when UIA tree traversal is blocked by a WinForms modal.
    /// </summary>
    private static IntPtr FindDialogHwndViaWin32()
    {
        IntPtr found = IntPtr.Zero;
        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (!NativeMethods.IsWindowVisible(hWnd))
                return true; // continue

            var sb = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
            if (sb.ToString() == "#32770")
            {
                found = hWnd;
                return false; // stop enumeration
            }
            return true; // continue
        }, IntPtr.Zero);
        return found;
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

    private static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
