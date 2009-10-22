Dim wApp
Set wApp = CreateObject("Word.Application")

Dim fso
Set fso = CreateObject("Scripting.FileSystemObjecT")

wApp.DisplayAlerts = 0

wApp.Documents.Open fso.GetFile("example.doc").Path, false, true

Dim currDoc
Set currDoc = wApp.Documents(1)

WScript.Echo "Word count: " & currDoc.Words.count

Dim word
For Each word In currDoc.Words
	WScript.Echo word
Next

currDoc.Close
wApp.Quit