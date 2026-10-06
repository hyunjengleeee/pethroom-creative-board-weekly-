param([string[]]$Files, [string]$Out, [int]$Tile = 520, [int]$Cols = 3)
Add-Type -AssemblyName System.Drawing
$rows = [math]::Ceiling($Files.Count / $Cols)
$bmp = New-Object System.Drawing.Bitmap ($Tile * $Cols), (($Tile + 30) * $rows)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::White)
$g.InterpolationMode = 'HighQualityBicubic'
$font = New-Object System.Drawing.Font 'Arial', 14, ([System.Drawing.FontStyle]::Bold)
for ($i = 0; $i -lt $Files.Count; $i++) {
  $img = [System.Drawing.Image]::FromFile($Files[$i])
  $x = ($i % $Cols) * $Tile; $y = [math]::Floor($i / $Cols) * ($Tile + 30)
  $g.DrawString([System.IO.Path]::GetFileNameWithoutExtension($Files[$i]), $font, [System.Drawing.Brushes]::Black, $x + 4, $y + 4)
  $g.DrawImage($img, $x, $y + 30, $Tile, $Tile)
  $img.Dispose()
}
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Jpeg)
$g.Dispose(); $bmp.Dispose()
