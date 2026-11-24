# Windows only: screenshot of this PC's screen (what the user sees right now).
# Usage: shot.cmd [out.png]   (prints the PNG path)
param([string]$Out = (Join-Path ([IO.Path]::GetTempPath()) 'word-shot.png'))
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public static class WordLiveDpi { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); }'
[WordLiveDpi]::SetProcessDPIAware() | Out-Null   # real pixels on a scaled (125%, 150%) display
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$screen = [Windows.Forms.SystemInformation]::VirtualScreen
$bmp = New-Object Drawing.Bitmap $screen.Width, $screen.Height
$g = [Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [Drawing.Point]::Empty, $screen.Size)
$Out = [IO.Path]::GetFullPath($Out)
$parent = Split-Path $Out
if (-not (Test-Path $parent)) { New-Item -ItemType Directory $parent | Out-Null }
$bmp.Save($Out, [Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
$Out
