Set objfso = createobject("scripting.filesystemobject")
Set file = objfso.createtextfile ( "H:\file.txt" , true )	 'create the file
file.writeline "This is the text I'm writing," 	'write a line
file.writeblanklines (2)				 'write 2 lines blank
file.writeline "Here more text"				'wrote another line of text
file. close 						'close the file