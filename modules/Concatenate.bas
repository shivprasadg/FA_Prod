Attribute VB_Name = "Concatenate"
Option Compare Database
Option Explicit

'> Hello,
'> I would Combining values from fields of more than 1 record in 1 field.
'> For example
'>
'> Name     firstname    number
'>
'> Heide    Marcel       1
'> Heide    Marcel       2
'> Heide    Marcel       3
'> Kooring  Yvonne       7
'> Kooring  Yvonne       14
'>
'> Number is a unique number
'>
'> Result should look like
'> Heide    Marcel       1,2,3
'> Kooring  Yvonne       7,14
'>

' The only way I know to do this is through code.  The attached VBA routine
' creates a table named "tblOriginal" and populates it with data (like above).
' Then it makes a copy of that table and aggregates the "number" field for
' matching "names".
'
' Works in A2k if you set a reference to Microsoft DAO Library.


Public Function FixTable(Column1 As String, Column2 As String, TableX As String) As Boolean
On Error Resume Next

Dim db As DAO.DataBase, RST As DAO.Recordset, SQLz As String
Dim strColumn1 As String, strColumn2 As String

Set db = CurrentDb()
'Call RecreateTables(db)
'Stop
SQLz = "SELECT " & Column1 & ", " & Column2 & " FROM " & TableX & " ORDER BY " & Column2 & ";"
Set RST = db.OpenRecordset(SQLz, dbOpenSnapshot)

If Not RST.BOF And Not RST.EOF Then
  RST.MoveFirst
  strColumn1 = RST(" & Column1 & ")
  strColumn2 = RST(" & Column2 & ")

  RST.MoveNext
  Do Until RST.EOF
    If strColumn1 = RST!Column1 Then
      strColumn2 = strColumn2 & ", " & RST!Column2
    Else
      SQLz = "INSERT INTO tblCopy (Column1, Column2)  " _
           & "VALUES('" & strColumn1 & "','" & strColumn2 & "')"
      db.Execute SQLz
      strColumn1 = RST!Column1
      strColumn2 = RST!Column2
    End If
    RST.MoveNext
  Loop

  ' Insert Last Record
  SQLz = "INSERT INTO tblCopy (Column1, Column2)  " _
       & "VALUES('" & strColumn1 & "','" & strColumn2 & "')"
  db.Execute SQLz
End If

Set RST = Nothing
Set db = Nothing

DoCmd.OpenForm "frmOutput"

End Function

Private Function RecreateTables(ByRef dbs As DAO.DataBase)
On Error Resume Next
Dim SQLz As String

' Delete Table, if exists
If DCount("*", "MsysObjects", "[Name]='tblOriginal'") = 1 Then
  DoCmd.DeleteObject acTable, "tblOriginal"
End If

SQLz = "CREATE TABLE tblOriginal (Column1 Text(10), Column2 Text(10))"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('A','1')"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('A','2')"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('B','1')"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('B','2')"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('B','3')"
dbs.Execute SQLz

SQLz = "INSERT INTO tblOriginal (Column1, Column2) VALUES ('C','1')"
dbs.Execute SQLz

' Delete Table, if exists
If DCount("*", "MsysObjects", "[Name]='tblCopy'") = 1 Then
    DoCmd.DeleteObject acTable, "tblCopy"
End If

' Create Temp Table
SQLz = "SELECT Column1, Column2 INTO tblCopy " _
       & "FROM tblOriginal WHERE 1 = 0;"
dbs.Execute SQLz

End Function
