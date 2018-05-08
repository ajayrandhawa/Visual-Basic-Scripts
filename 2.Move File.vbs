dim filesys 
set filesys=CreateObject("Scripting.FileSystemObject") 
If filesys.FileExists("D:\ajay.txt") Then 
filesys.MoveFile "D:\ajay.txt", "H:\" 
End If 
