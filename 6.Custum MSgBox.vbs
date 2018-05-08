Set objshell = createobject("Wscript.shell")
rmessage = objshell.popup ( "This is a test message" ,    1   , "Message Popup" ,16)
									'message			'time     'title		'icon
'Value  Description
'0 OK 
'1 OK and Cancel 
'2 Abort, Retry, and Ignor
'3 Yes, No, and Cancel 
'4 Yes and No 
'5 Retry and Cancel


'16 Stop / Error 
'32 Question 
'48 Exclamation 
'64 Information





