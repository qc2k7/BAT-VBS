param (
    [Parameter(Mandatory = $true)]
    [string]$WallpaperPath
)

# Validate the file exists
if (-not (Test-Path $WallpaperPath)) {
    Write-Error "Wallpaper file not found: $WallpaperPath"
    exit 1
}

# Convert to full path
$WallpaperFullPath = Resolve-Path $WallpaperPath

# Set wallpaper style values (0 = Center, 2 = Stretch, 6 = Fit, 10 = Fill)
$WallpaperStyle = "10"      # Fill
$TileWallpaper = "0"

# Set registry keys (cleaner method)
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper -Value $WallpaperFullPath
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value $WallpaperStyle
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value $TileWallpaper

# Apply the wallpaper using native Windows API
Add-Type @"
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

$SPI_SETDESKWALLPAPER = 20
$SPIF_UPDATEINIFILE = 0x01
$SPIF_SENDWININICHANGE = 0x02

[NativeMethods]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $WallpaperFullPath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDWININICHANGE)

Write-Output "Wallpaper set to: $WallpaperFullPath"
