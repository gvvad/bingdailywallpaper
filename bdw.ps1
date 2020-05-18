$HOST_URL = "http://bing.com"
$JSON_URL = "/HPImageArchive.aspx?format=js&idx=0&n=1"

$TEMP = $Env:TEMP
$TEMP_ITEM = Get-ChildItem -Path $TEMP

$TASK_NAME = "BingDailyWallpaper"
$SCRIPT_ITEM = Get-ChildItem -Path $MyInvocation.MyCommand.Path
$USER_NAME = "$Env:COMPUTERNAME\$Env:USERNAME"

Add-Type -AssemblyName System.Runtime.WindowsRuntime
function Await {
    param (
        $WinRtTask,
        $ResultType
    )
    
    foreach ($item in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
        if ($item.Name -eq 'AsTask' -and $item.GetParameters().Count -eq 1 -and $item.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1') {
            $asTaskGeneric = $item
            break
        }
    }
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $task = $asTask.Invoke($null, $WinRtTask)

    $task.Wait(-1) | Out-Null
    return $task.Result
}

function AwaitAction {
    param (
        $WinRtAction
    )
    
    foreach ($item in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
        if ($item.Name -eq 'AsTask' -and $item.GetParameters().Count -eq 1 -and !$item.IsGenericMethod) {
            $asTask = $item
            break
        }
    }
    $task = $asTask.Invoke($null, $WinRtAction)

    $task.Wait(-1) | Out-Null
    return $task.Result
}

function GetStorageFile {
    param (
        [string]$Path
    )
    [Windows.Storage.StorageFile, Windows.System.UserProfile.LockScreen, ContentType = WindowsRuntime] | Out-Null
    return Await -WinRtTask ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) -ResultType ([Windows.Storage.StorageFile])
}

function Get-LockscreenWallpaperPath {
    [Windows.System.UserProfile.LockScreen, Windows.System.UserProfile, ContentType = WindowsRuntime] | Out-Null
    return [Windows.System.UserProfile.LockScreen]::OriginalImageFile.AbsolutePath
}

function Set-LockScreenWallpaper {
    param (
        [string]$Path
    )
    $file = GetStorageFile -Path $Path
    AwaitAction -WinRtAction ([Windows.System.UserProfile.LockScreen]::SetImageFileAsync($file))
}

function Get-Wallpaper {
    return [Wallpaper]::GetWallpaper()
}

function Set-Wallpaper {
    param (
        [string]$Path,
        [ValidateSet('Tile', 'Center', 'Stretch', 'Fill', 'Fit', 'Span')]
        [string]$Style = 'Fill'
    )
    $StyleNum = @{
        Tile    = 0
        Center  = 1
        Stretch = 2
        Fill    = 3
        Fit     = 4
        Span    = 5
    }
    [Wallpaper]::SetWallpaper($Path, $StyleNum[$Style])
}

switch ($args[0]) {
    "set" {
        try {
            if (Get-ScheduledTask -TaskName $TASK_NAME -ErrorAction Ignore) {
                Write-Output "Task allready exist. Call with 'unset' parameter to remove task."
                exit 0
            }

            $trigger = @(New-ScheduledTaskTrigger -Daily -At 12:00)
            $trigger += New-ScheduledTaskTrigger -AtLogOn -User $USER_NAME

            $action = New-ScheduledTaskAction -Execute "$($SCRIPT_ITEM.DirectoryName)\nowin.exe" -Argument """powershell.exe"" ""-File $($SCRIPT_ITEM.FullName) -ExecutionPolicy bypass"""
            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 5 -RestartInterval 00:30
            Register-ScheduledTask -TaskName $TASK_NAME -Action $action -Trigger $trigger -Settings $settings
            
            Write-Output "Task created: $TASK_NAME"
        }
        catch {
            Write-Output "Task set error"
        }
        exit 0
    }

    "unset" {
        try {
            Unregister-ScheduledTask -TaskName $TASK_NAME
            Write-Output "Task removed"
        }
        catch {
            Write-Output "Task unset error"
        }
        exit 0
    }
}

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Win32;
    
    public class Wallpaper
    {
        public enum Style: int
        {
            Tile, Center, Stretch, Fill, Fit, Span, NoChange
        }
        private const uint MAX_PATH = 256;
    
        private const uint SPI_SETDESKWALLPAPER = 0x0014;
        private const uint SPI_GETDESKWALLPAPER = 0x0073;
    
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDCHANGE = 0x02;
    
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(uint uAction, uint uParam, string lpvParam, uint fuWinIni);
    
        public static void SetWallpaper(string path, Style style = Style.Fill)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
    
            int tile = 0;
            switch (style)
            {
                case Style.Tile:
                    key.SetValue(@"WallpaperStyle", "0");
                    tile = 1;
                    break;
                case Style.Center:
                    key.SetValue(@"WallpaperStyle", "0");
                    break;
                case Style.Stretch:
                    key.SetValue(@"WallpaperStyle", "2");
                    break;
                case Style.Fill:
                    key.SetValue(@"WallpaperStyle", "10");
                    break;
                case Style.Fit:
                    key.SetValue(@"WallpaperStyle", "6");
                    break;
                case Style.Span:
                    key.SetValue(@"WallpaperStyle", "22");
                    break;
                default:
                    break;
            }
            key.SetValue(@"TileWallpaper", tile.ToString());
            key.Close();
    
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0x0, path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }
    
        public static string GetWallpaper()
        {
            string buf = new string((char)0x0, (int)MAX_PATH);
    
            SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, buf, 0x0);
    
            int len = buf.IndexOf((char)0x0);
            return buf.Substring(0, len);
        }
    }
"@

$cur_wlp_item = Get-ChildItem -Path $(Get-Wallpaper)

Write-Output "Request provider update..."
try {
    $content = Invoke-WebRequest -Uri "$($HOST_URL+$JSON_URL)"
}
catch {
    Write-Output "Web invoke error: $($HOST_URL+$JSON_URL)"
    exit -1
}

$obj = ConvertFrom-Json -InputObject $content
$img_obj = $obj.images[0]

$file_path = Join-Path -Path $TEMP -ChildPath "$($img_obj.hsh).jpg"

if (-Not ($cur_wlp_item.FullName -eq $file_path)) {
    if (-Not (Test-Path -Path $file_path)) {
        Write-Output "Downloading file..."
        try {
            Invoke-WebRequest -Uri $($HOST_URL + $img_obj.url) -OutFile $file_path
        }
        catch {
            Write-Output "File request error!"
            exit -1
        }
        Write-Output "File: $file_path"
    }
    Write-Output "Set desktop wallpaper: $file_path"
    Set-Wallpaper -Path $file_path
        
    try {
        Write-Output "Set lockscreen wallpaper: $file_path"
        Set-LockScreenWallpaper -Path $file_path
    }
    catch { }

    if ($cur_wlp_item.Directory -eq $TEMP_ITEM) {
        Write-Output "Remove old file: $($cur_wlp_item.FullName)"
        Remove-Item -Path $cur_wlp_item.FullName
    }
}
