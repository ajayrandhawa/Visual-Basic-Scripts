Set objfso = createobject("scripting.filesystemobject")
Set file = objfso.getfile("H:\adf.pdf")	 'get control over PDF file
Msgbox file.attributes' message with the file ATTRIBUTES

'In this example we see that after declaring the object, what we do is get control of the tutorial.pdf file, for which we use getfile, and assign the file to the variable file. . Then and now with the file variable, we can use directly to display its attributes, which will be nothing but a number that will encompass all the attributes constants 
'constants that refer to the file attributes are:

'Value
'0 Normal 
'1 Read Only 
'2 Hidden 
'4 System
'8 drive letter 
'16 folder / directory
'32 File 
'64 Link or direct access
'128 Compressed

'As I mentioned before, attributes return a single value will be the sum of each of the values ​​for each attribute of the file. 
'For example: 
'A file that has attributes; read-only, hidden, system, and archive, will be worth 1 + 2 + 4 + 32 = 39 Change Attributes Set variable = objfso.getfile (route) variable.attributes = sumaatributos Example: