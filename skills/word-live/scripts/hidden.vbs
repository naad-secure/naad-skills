' Runs a PowerShell job with no visible window, then writes <out.txt>.done
' holding its exit code. prlctl does not wait for wscript, so collect.ps1
' waits on that marker instead.
' Usage: wscript //B hidden.vbs <job.ps1> <out.txt>
' A job just written on the Mac side can take a few seconds to appear in the
' shared folder, so wait for it before starting PowerShell.
Set args = WScript.Arguments
Set fs = CreateObject("Scripting.FileSystemObject")
For i = 1 To 100
  If fs.FileExists(args(0)) Then Exit For
  WScript.Sleep 100
Next
' prlctl can report a failed call that did start; the retry must not run the
' job a second time.
If fs.FileExists(args(1) & ".started") Then WScript.Quit 0
fs.CreateTextFile(args(1) & ".started", True).Close
cmd = "cmd /c powershell -NoProfile -ExecutionPolicy Bypass -File """ & args(0) & """ > """ & args(1) & """ 2>&1"
rc = CreateObject("WScript.Shell").Run(cmd, 0, True)
fs.CreateTextFile(args(1) & ".done", True).Write CStr(rc)
