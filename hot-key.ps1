Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class HotKeyManager
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }
"@

# Define your hotkey parameters
$hWnd = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
$hotkeyId = 1
$modifiers = 2  # Ctrl key
$virtualKey = 0x42  # 'B' key

# Register the hotkey
[HotKeyManager]::RegisterHotKey($hWnd, $hotkeyId, $modifiers, $virtualKey)

# Wait for the hotkey to be pressed
Write-Host "Press Ctrl + B to trigger the hotkey. Press Enter to exit."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Unregister the hotkey when done
[HotKeyManager]::UnregisterHotKey($hWnd, $hotkeyId)
