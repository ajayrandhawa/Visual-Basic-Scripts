Set objfso = createobject("scripting.filesystemobject")
Set file = objfso.getfile( "H:\adf.pdf" )	 'get control over PDF file
file.attributes = 34 'and hidden attribute will dearchivo