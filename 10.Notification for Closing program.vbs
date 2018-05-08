Set objshell = createobject("wscript.shell")
Set bloc = objshell.Exec("notepad")
Do while bloc.status = 0
	Wscript.Sleep 200	 'serves to pause for x milliseconds		
loop
msgbox "has closed the notebook"