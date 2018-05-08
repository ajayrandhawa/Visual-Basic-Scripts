Set objshell = createobject("wscript.shell")
Set bloc = objshell.exec ( "notepad" )		 'run notepad
wscript.Sleep 2000					 'wait two seconds
objshell.appactivate bloc.processid		 'we focus on the window notebook
wscript.sleep 200					'espera de milesimas
objshell.sendkeys "Tutorial vbs" 			'send a message sendkeys
objshell.sendkeys "{ENTER}" 			'then the previous message, ENTER
wscript.Sleep 2000					 'new waiting two seconds
Objshell.sendkeys "Testing the SendKeys function"  'send a second message line