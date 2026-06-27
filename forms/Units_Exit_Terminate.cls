Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Units_Exit_Terminate"
End Sub
Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.UnitID, 0) = 0 Or Nz(Me.UnitID, 0) <> UnitIDx Then
       Me.Form.RecordSource = "SELECT * FROM UnitExitOption " & _
                              "WHERE OptionTypeID =2 AND UnitId=" & UnitIDx & ";"
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
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub
Private Sub ModifyMe_Click()
If IsNull(Me.Parent.UnitID) Then Exit Sub
Call MakeExitOptionUnit(Me.Parent.UnitID, Me.Parent.SchID, Me.Parent.GroupRef, "Terminate")
End Sub
