Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'execute proc on every filename
Sub Traverse(path, proc)
	Dim baseFolder, currFolder, currFile
	
	Set baseFolder = fso.GetFolder(path)
	
	'execute proc on files
	For Each currFile In baseFolder.Files
		Execute proc & " " & chr(34) & currFile.path & chr(34)
	Next
	
	'recurse subdirs
	For Each currFolder In baseFolder.SubFolders
		Traverse currFolder.path, proc
	Next
	
End Sub

Traverse ".", "WScript.Echo"