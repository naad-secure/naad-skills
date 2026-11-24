# Windows side of word-live: runs a job (a PowerShell snippet) inside the Word
# that is open on this PC, with the helpers from prelude.ps1 ($word, $doc,
# Find-In, Fix-Bidi, Render, ...). Call it through word.cmd or render.cmd.
#
#   word.cmd job.ps1               run the snippet saved in job.ps1 (UTF-8)
#   word.cmd < job.ps1             same, snippet on stdin (UTF-8 bytes)
#   render.cmd out\page job.ps1    Render N in the job writes out\page-1.png ...
#
# Prefer a file: Windows PowerShell 5.1 pipes text to programs as ASCII, so
# Persian text piped in from a PowerShell prompt arrives as "????".
param(
  [Parameter(Position = 0)][string]$JobFile,
  [string]$RenderPrefix
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

if ($PSVersionTable.PSVersion.Major -ge 6) {
  # PowerShell 7 has no Marshal.GetActiveObject, so it cannot reach the running Word.
  [Console]::Error.WriteLine('word-live needs Windows PowerShell 5.1: run it through word.cmd / render.cmd')
  exit 2
}

if ($JobFile) {
  $snippet = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $JobFile).ProviderPath, [Text.Encoding]::UTF8)
} elseif ([Console]::IsInputRedirected) {
  $ms = New-Object IO.MemoryStream
  [Console]::OpenStandardInput().CopyTo($ms)
  $snippet = [Text.Encoding]::UTF8.GetString($ms.ToArray())
} else {
  [Console]::Error.WriteLine('usage: word.cmd job.ps1   |   render.cmd <prefix> job.ps1')
  exit 2
}
$snippet = $snippet.TrimStart([char]0xFEFF)

if ($RenderPrefix) {
  . (Join-Path $PSScriptRoot 'render-win.ps1')
  $WordLiveRender = @{ Prefix = [IO.Path]::GetFullPath($RenderPrefix); Count = 0 }
  $parent = Split-Path $WordLiveRender.Prefix
  if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory $parent | Out-Null }
}

# Prelude and snippet go into one temporary script, saved with a BOM so that
# PowerShell 5.1 reads the Persian text in it as UTF-8.
$prelude = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'prelude.ps1'), [Text.Encoding]::UTF8)
$job = Join-Path ([IO.Path]::GetTempPath()) "word-live-$PID.ps1"
[IO.File]::WriteAllText($job, $prelude + "`r`n" + $snippet + "`r`n", (New-Object Text.UTF8Encoding($true)))

$rc = 0
try {
  $global:LASTEXITCODE = 0
  & $job
  if ($LASTEXITCODE) { $rc = $LASTEXITCODE }
} catch {
  [Console]::Error.WriteLine(($_ | Out-String).Trim())
  $rc = 1
} finally {
  Remove-Item -LiteralPath $job -ErrorAction SilentlyContinue
}
# Environment.Exit, not exit: after Render used Windows.Data.Pdf, a normal
# shutdown of PowerShell 5.1 crashes and turns the exit code into -1073741819.
[Console]::Out.Flush()
[Environment]::Exit($rc)
