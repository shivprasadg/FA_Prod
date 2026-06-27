Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz As String

Private Sub ChildList_Click()
    Dim RST As DAO.Recordset
    Me.Form.FilterOn = False
    Set RST = Me.RecordsetClone
    RST.FindFirst "ChildPartID = " & Me.ChildList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End Sub

Private Sub ChildType_GotFocus()
    If Me.ChildClass & "test" = "test" Then
        MsgBox "You MUST Select a Class", , "No can do, without a Class"
        Me.ChildClass.SetFocus
    End If
End Sub

Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.UnitREF, 0) = 0 Or Nz(Me.UnitREF, 0) <> UnitIDx Then

        SQLz = "SELECT ChildPartID, ChildMake AS Make, ChildClass as Class, ChildType as [Type], childFinalCost AS Cost "
        SQLz = SQLz & "FROM ChildParts "
        SQLz = SQLz & "WHERE UnitREF= " & UnitIDx & "  ORDER BY ChildClass, ChildType;"
        Me.ChildList.RowSource = SQLz

        SQLz = "SELECT ChildParts.* "
        SQLz = SQLz & "FROM ChildParts "
        SQLz = SQLz & "WHERE UnitRef=" & UnitIDx & ";" ' ORDER BY cClass,  cClass;"
        Me.Form.RecordSource = SQLz

    End If
    If Me.Form.NewRecord = False And Nz(Me.UnitREF, 0) <> 0 Then Me.ChildPartSum = DSum("childFinalCost", "ChildParts", "[UnitREF]=" & UnitIDx & "")
    On Error GoTo 0
    Exit Function
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Function

Private Sub ChildClass_Click()
    If Me.ChildClass.Column(1) <> Nz(Me.ChildType.Column(1), 0) Then
       Me.ChildType = ""
       Me.cDesc = ""
    End If
    SQLz = "SELECT cPartType, cPartClassID " & _
           "FROM Lookup_PartTypes " & _
           "WHERE cPartClassID=" & Me.ChildClass.Column(1) & " " & _
           "ORDER BY cPartType;"
    Me.ChildType.RowSource = SQLz
End Sub

Private Sub Form_Current()
   On Error Resume Next
   If Me.Form.NewRecord = False And Nz(Me.UnitREF, 0) <> 0 Then Me.ChildPartSum = DSum("childFinalCost", "ChildParts", "[UnitREF]=" & [UnitREF] & "")
   On Error GoTo 0
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    DoCmd.RunCommand acCmdSaveRecord
End Sub
