Set objshell = createobject("wscript.shell")
Set Ellink = objshell.createshortcut( "H:\directo.lnk" )		 'create the link
Ellink.targetpath = "C:\windows\notepad.exe"		'completamos los valores
Ellink.windowstyle = 1
Ellink.hotkey = "CTRL+SHIFT+N"
Ellink.iconlocation = "C:\windows\notepad.exe,0"
Ellink.description = "Acceso directo a notepad"
Ellink.workingdirectory ="C:\"
Ellink.save 'we save the link