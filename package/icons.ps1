# Draws the Scriptless SharePoint icon ("{ }" crossed out on an indigo tile)
# and writes the images the app and package need:
#   -Ico      app.ico for the executables (16-256 px)
#   -Assets   the MSIX tiles at scale 100/200/400, and the taskbar and
#             Start sizes (targetsize-*, plated and unplated)
#   -McpIcon  the MCP bundle's icon.png (256 px)
# Windows only (System.Drawing).
param(
    [string]$Ico,
    [string]$Assets,
    [string]$McpIcon
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
# .NET resolves relative paths against the process's folder, not PowerShell's
function Resolve-Full([string]$p) { if ($p) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) } }
$Ico = Resolve-Full $Ico
$Assets = Resolve-Full $Assets
$McpIcon = Resolve-Full $McpIcon

function New-RoundedPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $d = 2 * $r
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

# Draws the mark into a square of side $s at ($x, $y).
function Draw-Mark([System.Drawing.Graphics]$g, [float]$x, [float]$y, [float]$s) {
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.InterpolationMode = 'HighQualityBicubic'
    $tile = New-RoundedPath $x $y $s $s ($s * 0.22)
    $brush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        [System.Drawing.PointF]::new($x, $y), [System.Drawing.PointF]::new($x + $s, $y + $s),
        [System.Drawing.Color]::FromArgb(255, 0x43, 0x38, 0xCA), [System.Drawing.Color]::FromArgb(255, 0x7C, 0x3A, 0xED))
    $g.FillPath($brush, $tile)

    # the braces
    $font = [System.Drawing.Font]::new('Segoe UI', [float]($s * 0.64), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = [System.Drawing.StringFormat]::new()
    $format.Alignment = 'Center'
    $format.LineAlignment = 'Center'
    $white = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $g.DrawString('{ }', $font, $white, [System.Drawing.RectangleF]::new($x - $s * 0.2, $y - $s * 0.05, $s * 1.4, $s), $format)

    # the strike: an amber bar from bottom left to top right
    $w = [Math]::Max(1.5, $s * 0.085)
    $amber = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 0xFB, 0xBF, 0x24), $w)
    $amber.StartCap = 'Round'; $amber.EndCap = 'Round'
    $a = [System.Drawing.PointF]::new($x + $s * 0.24, $y + $s * 0.76)
    $b = [System.Drawing.PointF]::new($x + $s * 0.76, $y + $s * 0.24)
    $g.DrawLine($amber, $a, $b)

    foreach ($o in $tile, $brush, $font, $format, $white, $amber) { $o.Dispose() }
}

# A transparent $width x $height image with the mark centred, $fill of the height.
function New-Image([int]$width, [int]$height, [double]$fill) {
    $bmp = [System.Drawing.Bitmap]::new($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $s = [Math]::Round([Math]::Min($width, $height) * $fill)
    Draw-Mark $g (($width - $s) / 2) (($height - $s) / 2) $s
    $g.Dispose()
    return $bmp
}

function Save-Png([System.Drawing.Bitmap]$bmp, [string]$path) {
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Get-PngBytes([int]$size) {
    $bmp = New-Image $size $size 1.0
    $ms = [System.IO.MemoryStream]::new()
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return , $ms.ToArray()
}

if ($Ico) {
    # an .ico of PNG images (Vista and later)
    $sizes = 16, 20, 24, 32, 40, 48, 64, 256
    [byte[][]]$images = foreach ($size in $sizes) { , (Get-PngBytes $size) }
    $ms = [System.IO.MemoryStream]::new()
    $w = [System.IO.BinaryWriter]::new($ms)
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)
    $offset = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $px = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
        $w.Write([byte]$px); $w.Write([byte]$px); $w.Write([byte]0); $w.Write([byte]0)
        $w.Write([uint16]1); $w.Write([uint16]32)
        $w.Write([uint32]$images[$i].Length); $w.Write([uint32]$offset)
        $offset += $images[$i].Length
    }
    foreach ($img in $images) { $w.Write($img) }
    $w.Flush()
    New-Item -ItemType Directory -Force (Split-Path $Ico) | Out-Null
    [System.IO.File]::WriteAllBytes($Ico, $ms.ToArray())
    Write-Host "  $Ico"
}

if ($Assets) {
    New-Item -ItemType Directory -Force $Assets | Out-Null
    # name, width, height (at scale 100), how much of the height the mark fills
    $tiles = @(
        @('StoreLogo', 50, 50, 1.0),
        @('Square44x44Logo', 44, 44, 1.0),
        @('Square150x150Logo', 150, 150, 0.6),
        @('Wide310x150Logo', 310, 150, 0.6),
        @('LargeTile', 310, 310, 0.5),
        @('SmallTile', 71, 71, 0.7)
    )
    foreach ($t in $tiles) {
        foreach ($scale in 100, 200, 400) {
            $f = $scale / 100
            Save-Png (New-Image ([int]($t[1] * $f)) ([int]($t[2] * $f)) $t[3]) (Join-Path $Assets "$($t[0]).scale-$scale.png")
        }
    }
    # taskbar, title bar and Start list sizes
    foreach ($size in 16, 20, 24, 30, 32, 36, 40, 48, 60, 64, 72, 80, 96, 256) {
        Save-Png (New-Image $size $size 1.0) (Join-Path $Assets "Square44x44Logo.targetsize-$size.png")
        Save-Png (New-Image $size $size 1.0) (Join-Path $Assets "Square44x44Logo.targetsize-${size}_altform-unplated.png")
        Save-Png (New-Image $size $size 1.0) (Join-Path $Assets "Square44x44Logo.targetsize-${size}_altform-lightunplated.png")
    }
    Write-Host "  $Assets"
}

if ($McpIcon) {
    New-Item -ItemType Directory -Force (Split-Path $McpIcon) | Out-Null
    Save-Png (New-Image 256 256 1.0) $McpIcon
    Write-Host "  $McpIcon"
}
