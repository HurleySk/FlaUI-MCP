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
/// Automate Windows file open/save dialogs.
/// Uses UIA for normal dialogs, falls back to pure Win32+SendInput
/// when WinForms modal dialogs block UIA.
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
            // Phase 1: Try UIA-based approach (fast path for non-modal dialogs)
            var uiaResult = await TryUiaApproach(path, action, timeout);
            if (uiaResult != null)
                return uiaResult;

            // Phase 2: Pure Win32+SendInput fallback (for WinForms modal dialogs
            // where all UIA calls hang because the parent UI thread is blocked)
            var win32Result = await TryWin32Approach(path, timeout);
            if (win32Result != null)
                return win32Result;

            return ErrorResult($"No file dialog appeared within {timeout}ms. Make sure you clicked a Browse/Open button first.");
        }
        catch (Exception ex)
        {
            return ErrorResult($"Failed to complete file dialog: {ex.Message}");
        }
    }

    /// <summary>
    /// UIA-based approach: find dialog via UIA tree, interact via automation patterns.
    /// Returns null if UIA times out (indicating modal blocking).
    /// </summary>
    private async Task<McpToolResult?> TryUiaApproach(string path, string action, int timeout)
    {
        var sw = Stopwatch.StartNew();
        // Give UIA at most half the timeout or 3s, whichever is less
        var uiaTimeout = Math.Min(timeout / 2, 3000);

        while (sw.ElapsedMilliseconds < uiaTimeout)
        {
            var findTask = Task.Run(() =>
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
                        catch { }
                    }
                }
                catch { }
                return null;
            });

            if (await Task.WhenAny(findTask, Task.Delay(2000)) == findTask)
            {
                var dialog = await findTask;
                if (dialog != null)
                {
                    await Task.Delay(200);
                    dialog.Focus();
                    await Task.Delay(100);
                    return await InteractViaUia(dialog, path, action);
                }
            }
            else
            {
                // UIA hung — modal is blocking, bail to Win32
                return null;
            }

            await Task.Delay(200);
        }

        return null; // UIA didn't find anything in time
    }

    /// <summary>
    /// Interact with a dialog found via UIA using automation patterns.
    /// </summary>
    private async Task<McpToolResult> InteractViaUia(Window dialog, string path, string action)
    {
        AutomationElement? filenameEdit = null;
        for (int attempt = 0; attempt < 3; attempt++)
        {
            filenameEdit = FindFilenameEdit(dialog);
            if (filenameEdit != null) break;
            await Task.Delay(300);
        }

        if (filenameEdit != null)
        {
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

            var confirmButton = FindConfirmButton(dialog, action);
            if (confirmButton != null)
            {
                if (confirmButton.Patterns.Invoke.IsSupported)
                    confirmButton.Patterns.Invoke.Pattern.Invoke();
                else
                    confirmButton.Click();

                return TextResult($"File dialog completed: {path}");
            }

            Keyboard.Press(VirtualKeyShort.ENTER);
            return TextResult($"File dialog completed (Enter key): {path}");
        }

        // UIA found the dialog but couldn't find controls — use keyboard
        dialog.Focus();
        await Task.Delay(100);
        Keyboard.Type(path);
        await Task.Delay(100);
        Keyboard.Press(VirtualKeyShort.ENTER);
        return TextResult($"File dialog completed (keyboard fallback): {path}");
    }

    /// <summary>
    /// Pure Win32+SendInput approach: find the dialog via EnumWindows (no UIA),
    /// focus it via SetForegroundWindow, and type the path via SendInput.
    /// This works even when WinForms modal dialogs block all UIA calls.
    /// </summary>
    private async Task<McpToolResult?> TryWin32Approach(string path, int timeout)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeout)
        {
            var hwnd = FindDialogHwndViaWin32();
            if (hwnd != IntPtr.Zero)
            {
                // Found dialog via Win32 — interact purely through Win32+SendInput
                await Task.Delay(300); // let dialog stabilize

                // Focus the dialog via Win32 (no UIA involved)
                NativeMethods.SetForegroundWindow(hwnd);
                await Task.Delay(200);

                // Try to find and set the filename edit via Win32 messages first
                var filenameSet = TrySetFilenameViaWin32(hwnd, path);
                if (filenameSet)
                {
                    await Task.Delay(100);
                    // Press Enter via SendInput to confirm
                    Keyboard.Press(VirtualKeyShort.ENTER);
                    return TextResult($"File dialog completed (Win32): {path}");
                }

                // Fallback: use SendInput keyboard — the filename combo typically
                // has focus when the dialog opens, so just type into it
                NativeMethods.SetForegroundWindow(hwnd);
                await Task.Delay(100);

                // Select all existing text, then type the new path
                Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                await Task.Delay(50);
                Keyboard.Type(path);
                await Task.Delay(200);
                Keyboard.Press(VirtualKeyShort.ENTER);

                return TextResult($"File dialog completed (Win32 keyboard): {path}");
            }

            await Task.Delay(200);
        }

        return null;
    }

    /// <summary>
    /// Try to set the filename in the dialog's edit control using Win32 messages.
    /// Finds the ComboBoxEx32 > ComboBox > Edit chain or a direct Edit child,
    /// then uses WM_SETTEXT to set the path.
    /// </summary>
    private static bool TrySetFilenameViaWin32(IntPtr dialogHwnd, string path)
    {
        // Standard file dialog has: ComboBoxEx32 (id 1148) > ComboBox > Edit
        var comboBoxEx = NativeMethods.FindWindowEx(dialogHwnd, IntPtr.Zero, "ComboBoxEx32", null);
        if (comboBoxEx != IntPtr.Zero)
        {
            var comboBox = NativeMethods.FindWindowEx(comboBoxEx, IntPtr.Zero, "ComboBox", null);
            if (comboBox != IntPtr.Zero)
            {
                var edit = NativeMethods.FindWindowEx(comboBox, IntPtr.Zero, "Edit", null);
                if (edit != IntPtr.Zero)
                {
                    NativeMethods.SendMessage(edit, NativeMethods.WM_SETTEXT, IntPtr.Zero, path);
                    return true;
                }
            }
            // Try direct Edit child of ComboBoxEx32
            var directEdit = NativeMethods.FindWindowEx(comboBoxEx, IntPtr.Zero, "Edit", null);
            if (directEdit != IntPtr.Zero)
            {
                NativeMethods.SendMessage(directEdit, NativeMethods.WM_SETTEXT, IntPtr.Zero, path);
                return true;
            }
        }

        // Fallback: look for any Edit control that's a direct child
        var editChild = NativeMethods.FindWindowEx(dialogHwnd, IntPtr.Zero, "Edit", null);
        if (editChild != IntPtr.Zero)
        {
            NativeMethods.SendMessage(editChild, NativeMethods.WM_SETTEXT, IntPtr.Zero, path);
            return true;
        }

        return false;
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
                return true;

            var sb = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
            if (sb.ToString() == "#32770")
            {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static AutomationElement? FindFilenameEdit(AutomationElement dialog)
    {
        try
        {
            var edits = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Edit));
            foreach (var edit in edits)
            {
                var autoId = edit.Properties.AutomationId.ValueOrDefault;
                if (autoId == "1148" || autoId == "1001")
                    return edit;
            }

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

            foreach (var btn in buttons)
            {
                var name = btn.Properties.Name.ValueOrDefault ?? "";
                if (name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    return btn;
            }

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
        public const int WM_SETTEXT = 0x000C;

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);
    }
}
