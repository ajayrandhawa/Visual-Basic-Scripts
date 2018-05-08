Set objfso = createobject("scripting.filesystemobject")
Set myfolder = objfso.getfolder ("H:\support" )		 'we obtain control over the folder
Set subfolders = myfolder.subfolders		 'get the collection of subfolders
For each s in subfolders				 'for each folder (s) in the collection (sub)
Msgbox s.name					 'message with the name
Next 