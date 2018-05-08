Set objshell = createobject("wscript.shell")
Msgbox objshell.regread("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures")				
	'It is all on one line, the key is too long