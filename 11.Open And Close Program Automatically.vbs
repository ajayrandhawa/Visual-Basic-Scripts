Set objshell = createobject("wscript.shell")
Set bloc = objshell.Exec("notepad")
Wscript.sleep 5000			'Pausa de 5000 milesimas, o 5 segundos		
bloc.terminate