Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String


Private Sub AssetID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.AssetID)
'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.AssetID
Me.Status.SetFocus
End Sub

Private Sub Buyer_DblClick(Cancel As Integer)
If Me.Buyer = "None" Then Exit Sub
    Call OpenBuyersFrom(Me.BuyerID, 0)
End Sub

Private Sub CompanyName_DblClick(Cancel As Integer)
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
    Me.Status.SetFocus
End Sub

Private Sub ExcelExport_Click()
FormatExcel "RemarketWrongStatusUnits", "C:\Faas\" & FaUserName & "\Temp\WrongStatus.xlsx", "wrong", , , , "Remaketing Wrong Status Units", True
End Sub

Private Sub Form_Current()
On Error GoTo FixMe
Me.txtCurrent = Me.AssetID
FixMe:
End Sub

Private Sub Status_DblClick(Cancel As Integer)
    Dim CommmitX As Variant
    Dim Message, NewStatus As String
    If Me.Status = "S" And Me.InvoicePaidDate & "test" = "test" Then
        Message = "This Unit should be set to 'R' status." & vbNewLine & "Do you want to set it?"
        NewStatus = "R"
    End If
    If Me.Status = "R" And Me.InvoicePaidDate & "test" <> "test" Then
        Message = "This Unit should be set to 'S' status." & vbNewLine & "Do you want to set it?"
        NewStatus = "S"
    End If
    If Me.Status = "T" And Me.BuyOutDate & "test" <> "test" Then
        Message = "This Unit should be set to 'R' status." & vbNewLine & "Do you want to set it?"
        NewStatus = "R"
    End If
    If NewStatus = "" Then Exit Sub
    CommmitX = MsgBox(Message, vbYesNo, "Modify")
    If CommmitX = vbNo Then Exit Sub
    SQLz = "UPDATE Units SET Units.UnitStatus ='" & NewStatus & "' WHERE UnitId=" & Me.AssetID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
End Sub
