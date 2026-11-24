# Prepended to every job: attach to the running Word and expose $word / $doc.
#
# Jobs work on Range objects and leave the user's cursor and scroll position
# alone, so the user can keep working in the document meanwhile. Set
# $Watch = $true at the top of a job to have Show select and scroll instead.
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

# While the user is typing or has a menu open, Word rejects COM calls with
# "Call was rejected by callee". This filter retries them for up to 30 s
# instead of failing the job.
Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
  [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int tick, IntPtr info);
  [PreserveSig] int RetryRejectedCall(IntPtr callee, int tick, int rejectType);
  [PreserveSig] int MessagePending(IntPtr callee, int tick, int pendingType);
}
public class RetryFilter : IOleMessageFilter {
  [DllImport("ole32.dll")] static extern int CoRegisterMessageFilter(IOleMessageFilter f, out IOleMessageFilter old);
  public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new RetryFilter(), out old); }
  public int HandleInComingCall(int callType, IntPtr caller, int tick, IntPtr info) { return 0; }
  public int RetryRejectedCall(IntPtr callee, int tick, int rejectType) { return (rejectType == 2 && tick < 30000) ? 200 : -1; }
  public int MessagePending(IntPtr callee, int tick, int pendingType) { return 2; }
}
"@
[RetryFilter]::Register()

$word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$doc = if ($word.Documents.Count -gt 0) { $word.ActiveDocument } else { $null }
$Watch = $false

# In watch mode, select a range and scroll it into view; otherwise do nothing.
function Show($range) { if ($Watch) { $range.Select(); $word.ActiveWindow.ScrollIntoView($range, $true) } }

# Run $action with $range selected (for Selection-only calls such as RtlRun),
# then put the user's selection and scroll position back.
function With-Selection($range, [scriptblock]$action) {
  $win = $word.ActiveWindow
  $sel = $word.Selection
  $saved = if ($sel.StoryType -eq 1) { $doc.Range($sel.Start, $sel.End) } else { $null }
  $scroll = $win.ActivePane.VerticalPercentScrolled
  $word.ScreenUpdating = $false
  try {
    $range.Select(); & $action
  } finally {
    if ($saved -and -not $Watch) { $saved.Select(); $win.ActivePane.VerticalPercentScrolled = $scroll }
    $word.ScreenUpdating = $true
  }
}

# Find $text inside $scope only (e.g. one table cell); returns the range or $null.
function Find-In($scope, $text) {
  $r = $doc.Range($scope.Start, $scope.End)
  $f = $r.Find; $f.ClearFormatting(); $f.MatchCase = $true; $f.Wrap = 0
  if ($f.Execute($text) -and $r.Start -ge $scope.Start -and $r.End -le $scope.End) { $r } else { $null }
}

# Find the first occurrence of $text (after $from, if given); returns the range or $null.
# Word's Find can jump back to an earlier match when the search range starts
# inside a table, so for repeated searches prefer Find-In on a known scope.
function Find-Text($text, $from = $null) {
  $r = if ($from) { $doc.Range($from.End, $doc.Content.End) } else { $doc.Content }
  $f = $r.Find; $f.ClearFormatting(); $f.MatchCase = $true; $f.Wrap = 0
  if ($f.Execute($text)) { $r } else { $null }
}

# Every occurrence of $text, collected before any of them is changed: once a
# match is replaced with tracking on, the old text stays as a deletion and a
# fresh search would find it again. Stops if Word's Find stops advancing.
function Find-All($text) {
  $hits = @(); $from = $null
  for ($i = 0; $i -lt 200; $i++) {
    $r = Find-Text $text $from
    if (-not $r -or ($from -and $r.Start -lt $from.End)) { break }
    $hits += $r; $from = $r
  }
  $hits
}

# Export pages of the document to PDF in the background (Word's own layout, so
# what it shows is exact). Lets the agent check an edit without scrolling the
# user's window. On Windows (render.cmd) the pages become PNGs right here; on
# the Mac the PDF goes back as base64 between markers and the Mac-side
# `render` script turns it into PNGs.
function Render([int]$from, [int]$to = $from) {
  # A plain .NET string: a path from Join-Path (a wrapped PSObject) makes Word
  # hang in ExportAsFixedFormat without an error.
  $pdf = [IO.Path]::Combine([IO.Path]::GetTempPath(), "word-live-render-$PID.pdf")
  # 17 = wdExportFormatPDF, 3 = wdExportFromTo, 0 = wdExportDocumentContent
  $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 3, $from, $to, 0)
  try {
    if ($WordLiveRender) { Convert-PdfToPng $pdf }
    else { "<<PDF>>" + [Convert]::ToBase64String([IO.File]::ReadAllBytes($pdf)) + "<<END>>" }
  } finally { Remove-Item $pdf }
}

# --- Editing helpers ------------------------------------------------------

# Filled-in answers are plain text: never bold, whatever the label next to them is.
# Latin text gets the document's usual Latin face; the Persian font is untouched.
function Set-Plain($range, $size = $null, $latinFont = 'Times New Roman') {
  $range.Font.Bold = 0; $range.Font.BoldBi = 0
  if ($latinFont) { $range.Font.Name = $latinFont }
  if ($size) { $range.Font.Size = $size; $range.Font.SizeBi = $size }
}

# Insert $text at the end of paragraph $para (before its mark); returns the new range.
function Append-Text($para, [string]$text) {
  $r = $para.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
  $start = $r.Start; $r.InsertAfter($text)
  $doc.Range($start, $start + $text.Length)
}

# Insert $text as a new plain paragraph (no bullet) after $para; returns its range.
# The break goes in before $para's own mark, because Word refuses
# InsertParagraphAfter on the last paragraph of a table cell, and the new
# paragraph is then found by position: after the break, $para.Next() can
# already point past the cell.
function Add-ParagraphAfter($para, [string]$text) {
  $indentL = $para.LeftIndent; $indentR = $para.RightIndent
  $end = $para.Range; $end.MoveEnd(1, -1) | Out-Null; $end.Collapse(0)
  $pos = $end.Start
  $end.InsertParagraphAfter()
  $p = $doc.Range($pos + 1, $pos + 1).Paragraphs(1)
  $p.Range.ListFormat.RemoveNumbers()
  $p.FirstLineIndent = 0; $p.LeftIndent = $indentL; $p.RightIndent = $indentR
  if ($text) { $doc.Range($pos + 1, $pos + 1).InsertAfter($text) }
  $doc.Range($pos + 1, $pos + 1 + $text.Length)
}

# Text put in with Range.Text or InsertAfter leaves spaces and punctuation next
# to Latin words as left-to-right, so in a Persian paragraph "(...)", ":" and
# "،" land on the wrong side. Mark every run of such characters that touches
# Persian text (or a paragraph edge) as right-to-left; ones between two Latin
# words stay as they are.
function Fix-Bidi($range) {
  $text = $range.Text
  $fa = '[؀-ۿﭐ-﷿ﹰ-﻿]'
  $spans = foreach ($m in [regex]::Matches($text, '[^\p{L}\p{N}\r\a]+')) {
    $before = if ($m.Index -gt 0) { $text[$m.Index - 1] } else { "`r" }
    $afterIdx = $m.Index + $m.Length
    $after = if ($afterIdx -lt $text.Length) { $text[$afterIdx] } else { "`r" }
    $edge = ($before -match "[\r\a]") -or ($after -match "[\r\a]")
    if ($edge -or ($before -match $fa) -or ($after -match $fa)) { ,@($m.Index, $m.Length) }
  }
  if (-not $spans) { return }
  With-Selection $range {
    foreach ($s in $spans) {
      $doc.Range($range.Start + $s[0], $range.Start + $s[0] + $s[1]).Select()
      $word.Selection.RtlRun()
    }
  }
}
