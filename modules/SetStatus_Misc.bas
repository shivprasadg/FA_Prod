Attribute VB_Name = "SetStatus_Misc"
 Sub testclientID()
MsgBox Forms!Unitsform!Text373
 End Sub

Sub SetStatus()
If Forms!Unitsform.AddUnits.Visible = True Then
If Forms!Unitsform.MLOrig = "Portfolio" Or Forms!Unitsform.MLOrig = "Other" Then
    Forms!Unitsform.UnitStatus.Value = "T"

    Else
    Forms!Unitsform.UnitStatus.Value = "B"
    End If



    If Forms!Unitsform.MLOrig = "Rmkt" And Forms!Unitsform.MLLessor = "FA LLC" Then
        Forms!Unitsform.UnitStatus.Value = "R"
    End If
        If Forms!Unitsform.MLOrig = "Rmkt" And (IsNull(Forms!Unitsform.MLLessor) Or Not Forms!Unitsform.MLLessor = "FA LLC") Then
        Forms!Unitsform.UnitStatus.Value = "K"
        End If

End If
End Sub

Sub stsnull()
'Forms!UnitsForm!RmktSubf!TitleOut.Value = "01/17/2012"

End Sub

Sub RenameTables()
Dim i As Integer
Dim dbRename As DataBase

    Dim strFullPath As String
    strFullPath = CurrentDb().Name

   Set dbRename = OpenDatabase(strFullPath)
   Dim strNewTableName As String


    dbRename.CreateTableDef

    For i = 0 To dbRename.TableDefs.Count - 1

      If Right(dbRename.TableDefs(i).Name, 1) = "1" Then

        strNewTableName = Left(dbRename.TableDefs(i).Name, Len(dbRename.TableDefs(i).Name) - 1)

        On Error Resume Next

        DoCmd.DeleteObject acTable, strNewTableName

        dbRename.TableDefs(i).Name = strNewTableName

      End If

    Next i

    dbRename.close
    Set dbRename = Nothing

End Sub

Public Function TableExists(sTable As String) As Boolean

    Dim db As DataBase
    Dim tbl As TableDef
    Set db = CurrentDb()

    TableExists = False

    For Each tbl In db.TableDefs
        If tbl.Name = sTable Then TableExists = True
    Next tbl

End Function
