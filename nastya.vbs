Dim WshShell, WshExec
Set WshShell = CreateObject("WScript.Shell")
do
    WScript.Sleep 500
    WshShell.SendKeys "Y"
	WScript.Sleep 500
    WshShell.SendKeys "F"
	WScript.Sleep 500
    WshShell.SendKeys "C"
	WScript.Sleep 500
    WshShell.SendKeys "N"
	WScript.Sleep 500
    WshShell.SendKeys "Z"
	WScript.Sleep 2000
    WshShell.SendKeys "{enter}"
loop
