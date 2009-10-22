Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim fileName
fileName = "example.mdb"

Dim conString, conStringProvider, conStringDataSource
conStringProvider = "PROVIDER=Microsoft.Jet.OleDb.4.0;"
conStringDataSource = "Data Source=" & fso.GetFile(fileName).Path & ";"
conString = conStringProvider & conStringDataSource

Dim con, res
Set con = CreateObject("ADODB.Connection")
Set res = CreateObject("ADODB.RecordSet")

con.Open(conString)
res.Open "SELECT Example FROM Example", con, 3, 3
res.AddNew
res("Example") = "Test"
res.Update
res.Close
con.Close