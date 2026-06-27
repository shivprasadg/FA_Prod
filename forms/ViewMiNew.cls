Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Public Function LoadData(UnitIDx As Long)
    On Error Resume Next
    If Nz(Me.UnitID, 0) = 0 Or Nz(Me.UnitID, 0) <> UnitIDx Then
        Me.Form.RecordSource = "SELECT vw_UMP_Mileage.* " & _
                               "FROM vw_UMP_Mileage Where UnitID=" & UnitIDx & " Order By LogActivityDate Desc;"
    End If
    On Error GoTo 0
End Function
Private Sub Form_Close()
 On Error GoTo Form_Close_Error
 On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_ViewMiNew"
   On Error GoTo 0
End Sub
