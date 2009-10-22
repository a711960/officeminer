Dim coll, wApp, cDoc, fso, fileName, paras, i

fileName = "example.doc"
Set fso = CreateObject("Scripting.FileSystemObject")

Set wApp = CreateObject("Word.Application")
wApp.DisplayAlerts = 0
wApp.Documents.Open fso.GetFile(fileName).path, false, true

Set paras = wApp.Documents(1).Paragraphs

For i = 1 To paras.count
	WScript.Echo "(" & i & ") " & paras.item(i).Range.Text
Next