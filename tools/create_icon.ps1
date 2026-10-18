# Deterministic Windows export of assets/tracomega.svg; no extra Python packages.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$projectRoot = if ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot } else { (Get-Location).Path }
$assetDirectory = Join-Path $projectRoot 'assets'
$bitmap = New-Object System.Drawing.Bitmap 256,256
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::Transparent)
$navy = [System.Drawing.ColorTranslator]::FromHtml('#142c48')
$teal = [System.Drawing.ColorTranslator]::FromHtml('#35d6c6')
$white = [System.Drawing.ColorTranslator]::FromHtml('#f3fbff')
function RoundedRect($drawing, $color, $x, $y, $width, $height, $radius) {
    $shape = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = 2 * $radius
    $shape.AddArc($x,$y,$diameter,$diameter,180,90)
    $shape.AddArc(($x+$width-$diameter),$y,$diameter,$diameter,270,90)
    $shape.AddArc(($x+$width-$diameter),($y+$height-$diameter),$diameter,$diameter,0,90)
    $shape.AddArc($x,($y+$height-$diameter),$diameter,$diameter,90,90)
    $shape.CloseFigure()
    $brush = New-Object System.Drawing.SolidBrush $color
    $drawing.FillPath($brush,$shape)
    $brush.Dispose(); $shape.Dispose()
}
RoundedRect $graphics $navy 12 12 232 232 54
$pen = New-Object System.Drawing.Pen $teal,15
$graphics.DrawEllipse($pen,46,46,130,130)
$pen.Dispose()
$pen = New-Object System.Drawing.Pen $teal,24
$pen.StartCap = $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawLine($pen,158,158,210,210)
$pen.Dispose()
$pen = New-Object System.Drawing.Pen $white,9
$pen.StartCap = $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawLine($pen,111,65,111,77)
$pen.Dispose()
$brush = New-Object System.Drawing.SolidBrush $teal
$graphics.FillEllipse($brush,104,55,14,14); $brush.Dispose()
RoundedRect $graphics $white 73 82 76 62 16
$brush = New-Object System.Drawing.SolidBrush $navy
$graphics.FillEllipse($brush,89,101,14,14)
$graphics.FillEllipse($brush,119,101,14,14); $brush.Dispose()
RoundedRect $graphics $navy 96 125 30 5 2.5
$graphics.Dispose()
$bitmap.Save((Join-Path $assetDirectory 'tracomega.png'),[System.Drawing.Imaging.ImageFormat]::Png)
$sizes = @(16,24,32,48,64,128,256)
$frames = [System.Collections.Generic.List[byte[]]]::new()
foreach ($size in $sizes) {
    $frame = New-Object System.Drawing.Bitmap $size,$size
    $drawing = [System.Drawing.Graphics]::FromImage($frame)
    $drawing.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $drawing.DrawImage($bitmap,0,0,$size,$size)
    $stream = New-Object System.IO.MemoryStream
    $frame.Save($stream,[System.Drawing.Imaging.ImageFormat]::Png)
    $frames.Add($stream.ToArray())
    $stream.Dispose(); $drawing.Dispose(); $frame.Dispose()
}
$outputStream = [System.IO.File]::Create((Join-Path $assetDirectory 'tracomega.ico'))
$writer = [System.IO.BinaryWriter]::new($outputStream)
$writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($index = 0; $index -lt $sizes.Count; $index++) {
    $dimension = if ($sizes[$index] -eq 256) { 0 } else { $sizes[$index] }
    $writer.Write([byte]$dimension); $writer.Write([byte]$dimension)
    $writer.Write([byte]0); $writer.Write([byte]0)
    $writer.Write([uint16]1); $writer.Write([uint16]32)
    $writer.Write([uint32]$frames[$index].Length); $writer.Write([uint32]$offset)
    $offset += $frames[$index].Length
}
foreach ($frameBytes in $frames) { $writer.Write([byte[]]$frameBytes) }
$writer.Dispose(); $bitmap.Dispose()
Write-Output 'SVG, PNG and multi-resolution ICO ready in assets/'
