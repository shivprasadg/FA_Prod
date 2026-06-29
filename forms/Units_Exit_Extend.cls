Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub DeleteExtn_Click()
    Dim iNum As Integer
    Committ = MsgBox("Delete the Extension that Starts On: " & Me.StartOrPurposedDate & vbNewLine & "Are you SURE?", vbYesNo + vbCritical, "DELETE!!!")
    If Committ = vbNo Then Exit Sub
    DoCmd.RunSQL ("Delete From UnitExitOption Where LeaseOptionID=" & Me.LeaseOptionID & ";")
    Me.ExtendListBox.Requery
    Me.Form.Requery
End Sub

Private Sub ExtendListBox_Click()
Dim RST As DAO.Recordset
    Set RST = Me.RecordsetClone
    RST.FindFirst "LeaseOptionID = " & Me.ExtendListBox.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    RST.close
    Set RST = Nothing
End Sub

Private Sub Form_Close()
   On Error GoTo Form_Close_Error
On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Units_Exit_Extend"
End Sub
Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.UnitID, 0) = 0 Or Nz(Me.UnitID, 0) <> UnitIDx Then
        Me.Form.RecordSource = "SELECT * FROM UnitExitOption " & _
                               "WHERE OptionTypeID >2 AND UnitId=" & UnitIDx & " " & _
                               "ORDER BY CDate(StartOrPurposedDate) DESC;"
        Call SetControls
    End If
    On Error GoTo 0
End Function
Private Sub SetControls()
On Error GoTo FixMe
If Nz(Me.DatePosted, 0) <> 0 Then
Me.ModifyMe.caption = "Edit Option"
Me.Detail.Visible = True
Else
Me.ModifyMe.caption = "Add Option"
Me.Detail.Visible = False
End If
Dim SQLx As String
SQLx = "SELECT LeaseOptionID, OptionSubType AS Type, Format([StartOrPurposedDate],'mm/dd/yy') AS Start, UnitExitOption.TermMonths AS Term, Format([EndOrActualDate],'mm/dd/yy') AS [End], Rent " & _
       "From UnitExitOption WHERE OptionTypeID > 2 And UnitExitOption.UnitID = " & Me.UnitID & " " & _
       "ORDER BY Cdate(Format([StartOrPurposedDate],'mm/dd/yy')) DESC;"
Me.ExtendListBox.RowSource = SQLx
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub

Private Sub ModifyMe_Click()
If IsNull(Me.Parent.UnitID) Then Exit Sub
Call MakeExitOptionUnit(Me.Parent.UnitID, Me.Parent.SchID, Me.Parent.GroupRef, "Extend")
End Sub
