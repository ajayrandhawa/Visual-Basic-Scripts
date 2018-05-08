Set objfso = createobject("scripting.filesystemobject")
Set archivotexto = objfso.opentextfile ( "H:\file.txt" , 1)	 'open the file
msgbox archivotexto.readline					 'read a line, the first
archivotexto.skipline							 'skipping a line
msgbox archivotexto.readline					 'read a line, the third
archivotexto. close 							'close the file