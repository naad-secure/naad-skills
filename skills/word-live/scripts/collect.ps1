# Runs as SYSTEM (off screen): waits for a hidden job to finish, prints its
# output as raw UTF-8 bytes and exits with the job's exit code. It deletes
# nothing, so a retried call prints the same output again; the Mac side
# removes the files once it has the output.
param([string]$Out)
$done = "$Out.done"
if (-not (Test-Path "$Out.started") -and -not (Test-Path $done)) {
  # Give a just-launched job a moment to start; past that, it never will.
  Start-Sleep -Seconds 5
  if (-not (Test-Path "$Out.started")) { Write-Error "job never started"; exit 125 }
}
$deadline = (Get-Date).AddMinutes(10)
while (-not (Test-Path $done)) {
  if ((Get-Date) -gt $deadline) { Write-Error "job timed out"; exit 124 }
  Start-Sleep -Milliseconds 100
}
$bytes = [IO.File]::ReadAllBytes($Out)
$stdout = [Console]::OpenStandardOutput(); $stdout.Write($bytes, 0, $bytes.Length); $stdout.Flush()
exit [int](Get-Content $done)
