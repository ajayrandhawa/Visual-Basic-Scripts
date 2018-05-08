Set objfso = createobject("scripting.filesystemobject")
Set archivotexto = objfso.opentextfile ( "H:\file.txt" , 8, true )			 'open the file
archivotexto.writeline "This is the text I'm writing," 			'write a line
archivotexto. close 