Option Explicit
Dim objShell
Set objShell = CreateObject("WScript.Shell")
Do
      objShell.SendKeys "%+{TAB}"
      Wscript.Sleep 500
Loop