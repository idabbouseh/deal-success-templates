$ErrorActionPreference = "Continue"
$SP = "C:\Users\I842506\AppData\Local\Temp\claude\C--Users-I842506-Downloads\389513d2-34fa-49ed-ab52-ca0f9cd4d127\scratchpad"
$DECKS = "$SP\decks2"
$RAW = "$SP\slides_raw2"
New-Item -ItemType Directory -Force $RAW | Out-Null

# decks worth a reader: everything except the ES-language duplicates and archive-ish extras
$skip = @("_2$", "customer_ppt_examples")
$files = Get-ChildItem "$DECKS\*.pptx" | Where-Object {
  $n = $_.BaseName
  -not ($skip | Where-Object { $n -match $_ })
}

Write-Host "Exporting $($files.Count) decks"
$ppt = New-Object -ComObject PowerPoint.Application
$done = 0
foreach ($f in $files) {
  $key = $f.BaseName
  $outdir = "$RAW\$key"
  if (Test-Path "$outdir\.complete") { $done++; continue }
  New-Item -ItemType Directory -Force $outdir | Out-Null
  try {
    $pres = $ppt.Presentations.Open($f.FullName, $true, $false, $false)
    $n = $pres.Slides.Count
    for ($i = 1; $i -le $n; $i++) {
      $dest = Join-Path $outdir ("{0:D3}.png" -f $i)
      if (-not (Test-Path $dest)) { $pres.Slides.Item($i).Export($dest, "PNG", 900, 506) }
    }
    $pres.Close()
    Set-Content "$outdir\.complete" $n
    $done++
    Write-Host "  [$done/$($files.Count)] $key : $n slides"
  } catch {
    Write-Host "  FAIL $key : $_"
    try { $pres.Close() } catch {}
  }
}
try { $ppt.Quit() } catch {}
Write-Host "EXPORT DONE: $done decks -> $RAW"
