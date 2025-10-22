Dim msg, sapi

Set WshShell = WScript.CreateObject("Wscript.Shell")

WshShell.Run "ttsSpam.vbs"
WshShell.Run "popupSpam.vbs"
WshShell.Run "keyboardSpam.vbs"