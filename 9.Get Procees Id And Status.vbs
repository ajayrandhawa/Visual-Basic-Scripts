Set objshell = createobject("wscript.shell")
Set bloc = Objshell.Exec("notepad")	
Msgbox bloc.status
Msgbox bloc.processid