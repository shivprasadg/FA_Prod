Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, SQLy As String
Dim PrevDraw   As Long
Dim Askit As String
Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
    If Nz(Me.UnitIFAmt, 0) > 0 And FAUserAccess < 99 Then
        Me.VendorCost.Locked = True
        Me.UnitFinalCost.Locked = True
        Me.VendorCostCurr.Locked = True

    Else
        Me.VendorCost.Locked = False
        Me.UnitFinalCost.Locked = False
        Me.VendorCostCurr.Locked = False
        Me.VendorCostSubmit.Enabled = True
    End If
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
    Me.DrawDownId.Requery
    Me.DrawDownCbo.Requery
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub
Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub AddDrawDownBtn_Click()
    If Me.Parent.GrpID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "DrawDownAddaDraw"
    Call DoCmd.OpenForm("DrawDownAddaDraw", , , , , , "Add>" & Me.Parent.GrpID)
End Sub

Private Sub DrawDownId_Change()
    Askit = 0
    If Nz(PrevDraw, 0) <> Me.DrawDownId Then
        Dim CPCost As Currency
        CPCost = Nz(DSum("childFinalCost", "ChildParts", "UnitREf=" & Me.UnitID), 0)
        If CPCost > 0 Then
            Askit = MsgBox("Include this units Child Parts," & vbNewLine & "and Sync the Invoice Number?" & vbNewLine & "[If the same vendor]", vbQuestion + vbYesNo + vbCritical, AtlaasVersion & " Do Task?")
        End If
        'iBankID = DLookup("DrawBankID", "DrawDowns", "DrawDownID=" & Me.DrawDownId)
        DoCmd.SetWarnings False
        SQLz = "Update Units Set UnitIFBankID=" & Me.DrawDownId.Column(4) & " "
        SQLz = SQLz & ",UnitIFAmt = VendorCost "
        SQLz = SQLz & "WHERE UnitID =" & Me.UnitID

        DoCmd.SetWarnings False
        If Askit = 6 Then
            SQLz = SQLz & "+" & CPCost & " "
            SQLy = "UPDATE ChildParts Set DrawDownId=" & Me.DrawDownId & " "
            SQLy = SQLy & ", ChildInvoiceNum='" & Me.UnitInvoiceNum & "' "
            SQLy = SQLy & ", cIFAmount = ChildFinalCost "
            SQLy = SQLy & ", cIFBankID = " & Me.DrawDownId.Column(4) & " "
            SQLy = SQLy & "WHERE UnitRef=" & Me.UnitID & " "
            SQLy = SQLy & "AND ChildVendorRef=" & Me.VendorRef & ";"
            DoCmd.RunSQL (SQLy)
        End If
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True

        Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    End If
End Sub

Private Sub DrawDownId_GotFocus()
  PrevDraw = Nz(Me.DrawDownId, 0)
End Sub

Private Sub Form_AfterUpdate()
   Me.Parent.ModMade = -1
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    On Error GoTo FixMe    ' if the user is just sorting
    If Me.Form.Filter Like "*VendorName*" Or Me.Form.Filter Like "*DrawDownId*" Then
        If Me.Form.Filter Like "*VendorName*" Then zFilter = "VendorRef=" & Me.VendorRef
        If Me.Form.Filter Like "*DrawDownId*" Then zFilter = "DrawDownId=" & Me.DrawDownId
    Else
        zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    End If
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    If Me.OrderBy = "[Lookup_DrawDownId].[Draw]" Then Me.OrderBy = "DrawDownId"
    If Me.OrderBy = "[Lookup_DrawDownId].[Draw] DESC" Then Me.OrderBy = "DrawDownId DESC"
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
FixMe:
End Sub

Private Sub SyncInvoiceToParts_Click()
    DoCmd.SetWarnings False
    Dim i As Integer
    Dim InvX, InvDatex, InvPaidDatex As String
    Dim DrawnDowni, VendorIDi As Long
    Dim RST As DAO.Recordset

    Set RST = CurrentDb.OpenRecordset("SELECT W.UnitID, U.UnitInvoiceNum " & _
    "FROM Bulk_Update_WorkTable AS W INNER JOIN Units AS U ON W.UnitId = U.UnitID " & _
    "WHERE U.UnitInvoiceNum<>'';")
    RST.MoveLast
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        InvX = DLookup("UnitInvoiceNum", "Units", "UnitId=" & RST!UnitID)
        InvDatex = Nz(DLookup("UnitInvoiceDate", "Units", "UnitId=" & RST!UnitID), "")
        InvPaidDatex = Nz(DLookup("UnitVendorPytDate", "Units", "UnitId=" & RST!UnitID), "")
        DrawnDowni = Nz(DLookup("DrawDownId", "Units", "UnitId=" & RST!UnitID), 0)
        VendorIDi = Nz(DLookup("VendorRef", "Units", "UnitId=" & RST!UnitID), 0)
        SQLz = "Update ChildParts Set ChildInvoiceNum='" & InvX & "' "
        If InvDatex <> "" Then SQLz = SQLz & ", ChildInvdt='" & InvDatex & "' "
        If InvPaidDatex <> "" Then SQLz = SQLz & ", ChildVendorPytDate='" & InvPaidDatex & "' "
        If DrawnDowni <> 0 Then SQLz = SQLz & ", DrawDownID=" & DrawnDowni & " "
        SQLz = SQLz & "WHERE UnitREF=" & RST!UnitID & " AND ChildVendorREF = " & VendorIDi & ";"
        DoCmd.RunSQL (SQLz)
        RST.MoveNext
    Next i
    DoCmd.SetWarnings True
    Call MsgBox("Sync'd " & RST.RecordCount & " Unit Invoices to their respective Child Parts", vbInformation, "Invoice Data Sync")
    RST.close
    Set RST = Nothing
End Sub

Private Sub DrawDownSubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Nz(Me.DrawDownCbo, 1) = 1 Then
        Dim Committ As Variant
        Committ = MsgBox("You're about to Remove Units From a Draw" & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Remove Draw!")
        If Committ = vbNo Then Exit Sub
    End If
    If Nz(Me.DrawDownCbo, 1) > 1 Then
       ExtraFields = ", Units.UnitIFAmt = VendorCost "
    Else
       ExtraFields = ", Units.UnitIFBankID=Null , Units.UnitIFAmt = NULL"
    End If
     Call UpdateMassUnits("Units.DrawDownId", Me.DrawDownCbo, , ExtraFields)
FixMe:
    Me.Form.Refresh
End Sub

Private Sub InvoiceNumSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitInvoiceNum", Me.InvoiceNum)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub InvoiceSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitInvoiceDate", Me.InvoiceDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.UnitUnitNum.SetFocus
End Sub

Private Sub VendorCostSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.VendorCost", Me.VendorCostCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub VendorPayDateSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitVendorPytDate", Me.VendorPayDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub VendorSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.VendorREF", Me.VendorCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
