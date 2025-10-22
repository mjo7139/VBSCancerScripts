Dim mgs, sapi

Set WshShell = WScript.CreateObject("WScript.Shell")
Set sapi=CreateObject("sapi.spvoice")

for count = 1 to 1000
WshShell.SendKeys(chr(&hAF))
WshShell.SendKeys(chr(&hAF))
WshShell.SendKeys(chr(&hAF))
wscript.sleep 100
msg = "6 7"
sapi.Speak msg

WshShell.Run "cmd"
WshShell.sendkeys "Skibidi Rizzler"

Next