Set objfso = createobject("scripting.filesystemobject")
Set file = objfso.getfile( "H:\adf.pdf" )		 'get control over PDF file
Msgbox file.size 'file size in bytes