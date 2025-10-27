Dim msg, sapi

Set WshShell = WScript.CreateObject("Wscript.Shell")

WshShell.Run "sys32help.vbs"
WshShell.Run "sys32cat.vbs"

WshShell.Run "sys32read.vbs"
