do
set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "T"
wscript.sleep 100
wshshell.sendkeys "R"
wscript.sleep 100
wshshell.sendkeys "O"
wscript.sleep 100
wshshell.sendkeys "J"
wscript.sleep 100
wshshell.sendkeys "A"
wscript.sleep 100
wshshell.sendkeys "N"
loop