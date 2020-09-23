Attribute VB_Name = "modDatabase"

Public db As New ADODB.Connection

Public Sub LoadDB()
On Error Resume Next

DB_Path = App.Path & "\Database\Main.mdb"

ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & DB_Path & ";DefaultDir=" & App.Path & ";UID=;PWD="

db.Close
db.ConnectionString = ConnectionString
db.ConnectionTimeout = 10
db.Open

End Sub
