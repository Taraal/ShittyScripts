Set wshShell = wscript.CreateObject("WScript.Shell") 
do
wscript.sleep 1800000
wshshell.sendkeys "%+{F4}"  
wshshell.sendkeys "{ENTER}"
loop