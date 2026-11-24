# Windows only (dot-sourced by word.ps1 for render.cmd): turns the PDF that
# Render exported into one PNG per page with Windows' own PDF renderer
# (Windows.Data.Pdf, built into Windows 10/11), and prints the PNG paths.
function Convert-PdfToPng([string]$pdf) {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

  $asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
  $asTaskOp = $asTask | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }
  $asTaskAction = $asTask | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' }
  $await = {
    param($op, [type]$type)
    $task = $asTaskOp.MakeGenericMethod($type).Invoke($null, @($op))
    $task.Wait() | Out-Null
    $task.Result
  }

  $file = & $await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
  $pdfDoc = & $await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
  for ($i = 0; $i -lt $pdfDoc.PageCount; $i++) {
    $page = $pdfDoc.GetPage($i)
    $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $options.DestinationWidth = 1600
    $options.DestinationHeight = [uint32][math]::Round(1600 * $page.Size.Height / $page.Size.Width)
    $buffer = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $asTaskAction.Invoke($null, @($page.RenderToStreamAsync($buffer, $options))).Wait() | Out-Null
    $WordLiveRender.Count++
    $png = "$($WordLiveRender.Prefix)-$($WordLiveRender.Count).png"
    $in = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($buffer.GetInputStreamAt(0))
    $out = [IO.File]::Create($png)
    try { $in.CopyTo($out) } finally { $out.Close(); $in.Dispose(); $buffer.Dispose(); $page.Dispose() }
    $png
  }
  # Note: once Windows.Data.Pdf has been used, PowerShell 5.1 crashes while
  # shutting down (exit code -1073741819); word.ps1 ends with Environment.Exit.
}
