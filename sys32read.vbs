Dim count
Set key = CreateObject("wscript.shell")

For count = 1 to 5000
key.sendkeys "Go"
wscript.sleep 1000
key.sendkeys "Panthers!"
wscript.sleep 1000
Next

