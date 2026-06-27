Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub ChildList_Click()
    Dim Childx As String
    Childx = Replace(Me.ChildList, "'", "''")
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From Bulk_Update_WorkTable_Child;")
    DoCmd.SetWarnings True
    SQLz = "INSERT INTO Bulk_Update_WorkTable_Child ( UnitID, ChildPartID, UnitUnitNum, UnitVIN) " & _
           "SELECT DISTINCTROW UnitRef, ChildPartID, UnitUnitNum, UnitVIN From Bulk_Update_SelectedUnits_Child Where UID ='" & Childx & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = "Select UnitUnitNum, UnitVIN, * From ChildParts as CP Inner Join Bulk_Update_WorkTable_Child As CWT On CWT.ChildPartID= CP.ChildPartID;"
    Me.Form.RecordSource = SQLz
    Me.Detail.Visible = True
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub cIFAmount_AfterUpdate()
    If Me.cIFAmount > Me.childFinalCost Then
        Call MsgBox("Sorry, you cannot enter a Value greater then the Part Cost." & vbNewLine & "This will be undone", vbCritical, "Stop spending money you don't have!")
        Me.Undo
    End If
End Sub
Private Sub cIFBankId_AfterUpdate()
    If Me.DrawDownId & "test" <> "test" Then Call SetInterimBank(Me.DrawDownId, Nz(Me.cIFBankId, 0))
End Sub
Private Sub cIFDate_AfterUpdate()
    If Me.DrawDownId & "test" <> "test" Then Call SetInterimFundDates(Me.DrawDownId, Nz(Me.cIFDate, ""))
End Sub

Private Sub DrawDownId_AfterUpdate()
    If Me.DrawDownId & "test" <> "test" Then
        Me.cIFAmount = Me.childFinalCost
    Else
        Me.cIFAmount = Null
    End If
End Sub

Private Sub Form_DataSetChange()
    Me.DrawDownCbo.Requery
    Me.DrawDownId.Requery
End Sub

Private Sub MakeExcelExport_Click()
    SQLz = "SELECT W.UnitID As AssetID"                             'A
    SQLz = SQLz & ", CP.ChildMake"                                  'B
    SQLz = SQLz & ", CP.ChildClass"                                 'C
    SQLz = SQLz & ", CP.ChildType"                                  'D
    SQLz = SQLz & ", CP.ChildVIN AS [Child SerialNum]"              'E
    SQLz = SQLz & ", Vendors.VendorSNm as Vendor "                  'F
    SQLz = SQLz & ", CP.childFinalCost AS [Child Cost]"             'G
    SQLz = SQLz & ", CP.ChildInvoiceNum AS [Child InvNum]"          'H
    SQLz = SQLz & ", CP.ChildInvdt AS [Child Inv Date]"             'I
    SQLz = SQLz & ", CP.ChildVendorPytDate AS [Paid Date]"          'J
    SQLz = SQLz & "FROM (Bulk_Update_WorkTable as W "
    SQLz = SQLz & "INNER JOIN ChildParts AS CP ON W.UnitID = CP.UnitREF) "
    SQLz = SQLz & "INNER JOIN Vendors ON CP.ChildVendorREF = Vendors.VendorID "
    SQLz = SQLz & "ORDER BY W.UnitID, CP.ChildMake, CP.ChildClass; "

    Call SendReport2Excel(SQLz, "Child Parts", "Child Parts", True, "|A,BC|G,$E|I,TC|J,TC|", , True)

End Sub

Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.UnitID
    Call ClearBoxes
FixMe:
    Exit Sub
End Sub
Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
    Me.DrawDownCbo.Requery
    Me.DrawDownId.Requery
End Sub



Private Sub Form_Load()
    Call ClearBoxes
End Sub

Private Sub SYNCChild2Units_Click()
    SQLz = "INSERT INTO Bulk_Update_WorkTable ( UnitID ) " & _
        "SELECT Bulk_Update_WorkTable_Child.UnitID FROM Bulk_Update_WorkTable_Child;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From Bulk_Update_WorkTable")
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call Me.Parent.ReQueryAll
    Me.Form.Refresh
End Sub

Private Sub ClassCbo_Click()
    If Me.ClassCbo.Column(1) <> Nz(Me.TypeCbo.Column(1), 0) Then
       Me.TypeCbo = ""
       Me.cDescCbo = ""
    End If
    SQLz = "SELECT cPartType, cPartClassID " & _
           "FROM Lookup_PartTypes " & _
           "WHERE cPartClassID=" & Me.ClassCbo.Column(1) & " " & _
           "Order By cPartType;"
    Me.TypeCbo.RowSource = SQLz
End Sub

Private Sub ChildClass_Click()
    If Me.ChildClass.Column(1) <> Nz(Me.ChildType.Column(1), 0) Then
       Me.ChildType = ""
       Me.cDesc = ""
    End If
    SQLz = "SELECT cPartType, cPartClassID " & _
           "FROM Lookup_PartTypes " & _
           "WHERE cPartClassID=" & Me.ChildClass.Column(1) & " " & _
           "Order By cPartType;"
    Me.ChildType.RowSource = SQLz
End Sub

Private Sub TypeCbo_GotFocus()
    If Me.ClassCbo & "test" = "test" Then
        MsgBox "You MUST Select a Class", , "No can do, without a Class"
        Me.ClassCbo.SetFocus
    End If
End Sub

Private Sub ChildType_GotFocus()
    If Me.ChildClass & "test" = "test" Then
        MsgBox "You MUST Select a Class", , "No can do, without a Class"
        Me.ChildClass.SetFocus
    End If
End Sub

Private Sub MakeSubmit_Click()
    If Nz(Me.MakeCbo, "") = "" Then Exit Sub
    If MassConfirmIT("Child Make", Me.MakeCbo) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildMake", Me.MakeCbo)
    'SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildMake = '" & Me.MakeCbo & "' " & _
        "WHERE UID= '" & Me.ChildList & "';"
     'SQLz = "UPDATE ChildParts  SET ChildMake = '" & Me.MakeCbo & "' " & _
             "WHERE ChildPartID IN (Select ChildPart From Bulk_Update_WorkTable_Child);"
     ' DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub ClassSubmit_Click()
    If Nz(Me.ClassCbo, "") = "" Then Exit Sub
    If MassConfirmIT("Child Class", Me.ClassCbo) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildClass", Me.ClassCbo)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub TypeSubmit_Click()
    If Nz(Me.TypeCbo, "") = "" Then Exit Sub
    If MassConfirmIT("Child Type", Me.TypeCbo) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildType", Me.TypeCbo)
 '   SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildType = '" & Me.TypeCbo & "' " & _
        "WHERE UID= '" & Me.ChildList & "';"
 '   DoCmd.SetWarnings False
 '   DoCmd.RunSQL (SQLz)
 '   ModMade = -1
 '   DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub cDescSubmit_Click()
    If MassConfirmIT("Child Desc", Me.cDescCbo) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.cDesc", Me.cDescCbo)
 '   SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET cDesc = '" & Me.cDescCbo & "' " & _
 '       "WHERE UID= '" & Me.ChildList & "';"
 '   DoCmd.SetWarnings False
 '   DoCmd.RunSQL (SQLz)
 '   ModMade = -1
 '   DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub VendorSubmit_Click()
    If Nz(Me.ChildVendorCbo.Column(1), "") = "" Then Exit Sub
    If MassConfirmIT("Child Vendor", Me.ChildVendorCbo.Column(1)) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildVendorREF", Me.ChildVendorREF)
'   SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildVendorREF = '" & Me.ChildVendorREF & "' " & _
        "WHERE UID= '" & Me.ChildList & "';"
'   DoCmd.RunSQL (SQLz)
'   ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub InvoiceNumSubmit_Click()
    If MassConfirmIT("Invoice Num", Me.InvoiceNum) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildInvoiceNum", Me.InvoiceNum)
'   SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildInvoiceNum = '" & Me.InvoiceNum & "' " & _
'       "WHERE UID= '" & Me.ChildList & "';"
'   DoCmd.RunSQL (SQLz)
'   ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub InvoiceDateSubmit_Click()
    If MassConfirmIT("Invoice Date", Me.InvoiceDate) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildInvdt", Me.InvoiceDate)
 '   SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildInvdt = #" & Me.InvoiceDate & "# " & _
 '       "WHERE UID= '" & Me.ChildList & "';"
 '   DoCmd.RunSQL (SQLz)
 '   ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub VendorPayDateSubmit_Click()
    If MassConfirmIT("Vendor Paid Date", Me.VendorPayDate) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildVendorPytDate", Me.VendorPayDate)
'    SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET ChildVendorPytDate = #" & Me.VendorPayDate & "# " & _
'        "WHERE UID= '" & Me.ChildList & "';"
'    DoCmd.RunSQL (SQLz)
'    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub CostSubmit_Click()
    If Me.CostCurr & "test" = "test" Then Exit Sub
    If MassConfirmIT("Part Cost", Me.CostCurr) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.childFinalCost", Me.CostCurr)
'    SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET childFinalCost = " & Me.CostCurr & " " & _
'        "WHERE UID= '" & Me.ChildList & "';"
'    DoCmd.RunSQL (SQLz)
'    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub cIFDateSubmit_Click()
    If MassConfirmIT("IF Date", Me.cIFDateCBO) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.cIFDate", Me.cIFDateCBO)
'    SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child " & _
           "SET cIFDate = '" & Me.cIFDateCBO & "' " & _
           "WHERE UID= '" & Me.ChildList & "';"
'    DoCmd.RunSQL (SQLz)
'    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub cIfAmountSubmit_Click()
    If MassConfirmIT("IF Amount", "to match the Part Cost") = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE ChildParts  SET cIFAmount = childFinalCost " & _
           "WHERE ChildPartID IN (Select ChildPart From Bulk_Update_WorkTable_Child);"
    DoCmd.RunSQL (SQLz)
'    SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child "
'    SQLz = SQLz & " SET cIFAmount = childFinalCost "
'    SQLz = SQLz & " WHERE UID= '" & Me.ChildList & "';"
'    DoCmd.RunSQL (SQLz)
'    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub cIfBankSubmit_Click()
    If MassConfirmIT("IF Bank", Nz(Me.cIFBankCBO.Column(1), "")) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.cIFBankID", Me.cIFBankCBO)
  '  SQLz = "UPDATE DISTINCTROW Bulk_Update_SelectedUnits_Child SET cIFBankID = " & Me.cIFBankCBO & " " & _
  '      "WHERE UID= '" & Me.ChildList & "';"
  '  DoCmd.RunSQL (SQLz)
  '  ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub DrawDownSubmit_Click()
    If MassConfirmIT("Draw Number", Nz(Me.DrawDownCbo.Column(1), "")) = vbNo Then Exit Sub
    On Error GoTo FixMe
    If Me.DrawDownCbo & "test" = "test" Then
        'Removes The DrawID
        SQLz = "UPDATE ChildParts  SET " & _
        " DrawDownId = childFinalCost " & _
        ",cIFAmount = NULL " & _
        ",cIFDate = NULL " & _
        "WHERE ChildPartID IN (Select ChildPart From Bulk_Update_WorkTable_Child);"
    Else
        SQLz = "UPDATE ChildParts  SET " & _
        "DrawDownId = " & Me.DrawDownCbo & " " & _
        ",cIFAmount = ChildFinalCost " & _
        ",cIFDate = '" & Me.ChildList & "' " & _
        "WHERE ChildPartID IN (Select ChildPart From Bulk_Update_WorkTable_Child);"

    DoCmd.RunSQL (SQLz)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub ExcludeCostDocsReportsSubmit_Click()
    If MassConfirmIT("Exclude Docs Rpt", Me.ExcludeCostDocsReportsCbo) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ExcludeCostDocsReports", Me.ExcludeCostDocsReportsCbo)
  '  SQLz = "UPDATE ChildParts INNER JOIN Bulk_Update_WorkTable_Child ON ChildParts.ChildPartID = Bulk_Update_WorkTable_Child.ChildPartID " & _
  '      "SET ExcludeCostDocsReports = " & Me.ExcludeCostDocsReportsCbo & ";"
  '  DoCmd.RunSQL (SQLz)
  '  ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub EXCofaSubmit_Click()
    If MassConfirmIT("Exclude CofA", Me.EXCofaCbox) = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassPARTS("ChildParts.ChildEXCofa", Me.EXCofaCbox)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
