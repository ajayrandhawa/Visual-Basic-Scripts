Set objshell = createobject("wscript.shell")
Set ping = objshell.exec("cmd /c ping 127.0.0.1")
Msgbox ping.stdout.readall