Sub openURLinFirefox(strLink)
     Set wshShell = CreateObject("WScript.Shell")
     strFirefoxPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
         wshShell.Run strFirefoxPath & "https://twitter.com/" & strLink, 1, false
End Sub