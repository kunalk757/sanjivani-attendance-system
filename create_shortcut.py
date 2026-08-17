"""
Create/Refresh Windows Desktop and Start Menu shortcuts for Sanjivani Attendance
with explicit multi-resolution icon location and icon cache refresh.
"""

import os
import sys
import ctypes

def create_windows_shortcuts():
    try:
        import win32com.client
    except ImportError:
        # Fallback to subprocess powershell if pywin32 not installed
        pass

    base_dir = os.path.abspath(os.path.dirname(__file__))
    exe_path = os.path.join(base_dir, "dist", "Sanjivani Attendance.exe")
    ico_path = os.path.join(base_dir, "sanjivani.ico")

    if not os.path.exists(exe_path):
        print(f"[ERROR] Executable not found at: {exe_path}")
        return False

    if not os.path.exists(ico_path):
        ico_path = exe_path

    # PowerShell script to create shortcuts via WScript.Shell
    ps_script = f"""
    $WshShell = New-Object -ComObject WScript.Shell
    
    # 1. Desktop Shortcut
    $desktopFolders = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('CommonDesktopDirectory'),
        'C:\\Users\\kunal\\OneDrive\\Desktop',
        'C:\\Users\\kunal\\OneDrive\\Documents\\Desktop'
    )
    
    foreach ($d in $desktopFolders) {{
        if (Test-Path $d) {{
            # Remove old/stale shortcuts to prevent generic icon caching
            Get-ChildItem -Path $d -Filter '*Sanjivani*.lnk' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $d -Filter 'Face Attendance*.lnk' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            
            $lnkPath = Join-Path $d 'Sanjivani Attendance.lnk'
            $sc = $WshShell.CreateShortcut($lnkPath)
            $sc.TargetPath = '{exe_path}'
            $sc.WorkingDirectory = '{base_dir}'
            $sc.IconLocation = '{ico_path},0'
            $sc.Description = 'Sanjivani AI Face Attendance System'
            $sc.Save()
            Write-Host "Created Desktop Shortcut: $lnkPath"
            Write-Host "  -> TargetPath: $($sc.TargetPath)"
            Write-Host "  -> IconLocation: $($sc.IconLocation)"
            break
        }}
    }}
    
    # 2. Refresh Windows Shell Icon Cache
    try {{
        $code = @'
        using System;
        using System.Runtime.InteropServices;
        public class ShellIconCache {{
            [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        }}
'@
        Add-Type -TypeDefinition $code
        [ShellIconCache]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
        Write-Host "Windows Shell Icon Cache notified successfully."
    }} catch {{
        Write-Host "Shell notification note: $_"
    }}
    """
    
    import subprocess
    result = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
    return result.returncode == 0

if __name__ == "__main__":
    create_windows_shortcuts()
