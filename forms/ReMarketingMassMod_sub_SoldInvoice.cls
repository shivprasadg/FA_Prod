Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLset As String
Public RST    As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.UnitID
    SQLz = "Select ClientShortname As Client, Schedule, UnitGroup From vw_SixKeys WHERE AssetID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.Parent.BuyNo = "Client: " & RST!Client & " Sch|Grp: " & RST!Schedule & "|" & RST!UnitGroup
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Sub
Private Sub InvoiceRequestDate_AfterUpdate()
    Call CheckSoldStatus
End Sub
Private Sub InvoicePaidDate_AfterUpdate()
    Call CheckSoldStatus
End Sub
Private Sub InvoiceSentDate_AfterUpdate()
    Call CheckSoldStatus
End Sub
Private Sub BuyerRefSubmit_Click()
    On Error GoTo FixMe
    If Me.BuyerRefCBO & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.BuyerRef", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.BuyerRef", Me.BuyerRefCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub CheckSoldStatus()
    If IsDate(Me.InvoiceSentDate) And Not IsDate(Me.InvoicePaidDate) Then Me.UnitStatus = "PS"
    If IsDate(Me.InvoiceSentDate) And IsDate(Me.InvoicePaidDate) Then Me.UnitStatus = "S"
    If Not IsDate(Me.InvoiceSentDate) And Not IsDate(Me.InvoicePaidDate) And IsDate(BuyOutDate) Then Me.UnitStatus = "EZ"
End Sub

Private Sub RmktSPRAdjSubmit_Click()
    If Me.RmktSPRAdjCBO & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.RmktSPRAdj", RmktSPRAdjCBO, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub RmktDepositSubmit_Click()
    On Error GoTo FixMe
    If Me.RmktDepositCbo & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.Deposit", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.Deposit", Me.RmktDepositCbo, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub RmktDepdtSubmit_Click()
    On Error GoTo FixMe
    If Me.RmktDepositCbo & "test" = "test" Then
        Call UpdateRMKTtbl("Rmkt.RmktDepdt", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("Rmkt.RmktDepdt", Me.RmktDepdtCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub RmkSmSubmit_Click()
    On Error GoTo FixMe
    If Me.RmkSmCBO & "test" = "test" Then
        Call UpdateRMKTtbl("Rmkt.RmkSm", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("Rmkt.RmkSm", Me.RmkSmCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub SoldPriceSubmit_Click()
    On Error GoTo FixMe
    If Me.SoldPriceCbo & "test" = "test" Then
        Call UpdateRMKTtbl("Rmkt.SoldPrice", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("Rmkt.SoldPrice", Me.SoldPriceCbo, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub UnitStatusSubmit_Click()
    If Me.UnitStatusCbo & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRUnitTbl("Units.UnitStatus", Me.UnitStatusCbo, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub WholesaleSubmit_Click()
    On Error GoTo FixMe
    If Me.WholesaleCBO & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.Wholesale", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.Wholesale", Me.WholesaleCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub InvoiceRequestDateSubmit_Click()
    On Error GoTo FixMe
    If Me.InvoiceRequestDateCBO & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.InvoiceRequestDate", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.InvoiceRequestDate", Me.InvoiceRequestDateCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub InvoiceSentDateSubmit_Click()
    On Error GoTo FixMe
    If Me.InvoiceSentDateCBO & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.InvoiceSentDate", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.InvoiceSentDate", Me.InvoiceSentDateCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub InvoicePaidDateSubmit_Click()
    On Error GoTo FixMe
    If Me.InvoicePaidDateCBO & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.InvoicePaidDate", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.InvoicePaidDate", Me.InvoicePaidDateCBO, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub RmktBOSSentSubmit_Click()
    On Error GoTo FixMe
    If Me.RmktBOSSentCbo & "test" = "test" Then
        Call UpdateRMKTtbl("RMKT.RmktBOSSent", Null, , , Me.Parent.Name)
    Else
        Call UpdateRMKTtbl("RMKT.RmktBOSSent", Me.RmktBOSSentCbo, , , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Forms!Unitsform.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub
