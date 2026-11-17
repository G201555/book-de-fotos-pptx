param([string]$arq, [string]$saida, [string]$slides)
$pp = New-Object -ComObject PowerPoint.Application
try {
  $p = $pp.Presentations.Open($arq, $true, $false, $false)
  foreach ($n in $slides.Split(",")) { $p.Slides.Item([int]$n).Export("$saida\s$n.png", "PNG", 1280, 720) }
  $p.Close()
} finally { $pp.Quit() }
"ok"
