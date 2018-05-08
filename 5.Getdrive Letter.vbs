Set objfso = createobject("scripting.filesystemobject")
Set discs = objfso.drives					 'get record collection
For  each d in discs					 'for each disk (d) in the collection (disks)
Msgbox d.driveletter					 'message drive letter
Next 