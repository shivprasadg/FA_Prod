Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String
Dim Response   As String
Private Sub Form_Current()
   Me.ItemList.Requery
End Sub
Private Sub AddNew_Click()
    If Me.Dirty Then
        MsgBox "You Cannot Add a record at this time." & vbNewLine & "You Must SAVE/UPDATE This One First", vbCritical, "Editing Data!"
        Exit Sub
        End
    End If
        Response = MsgBox("You are about to Add a WO Item." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Add WO Item?")
        If Response = vbYes Then
            DoCmd.GotoRecord , , acNewRec
            Me.WorkOrderID = Me.Parent.WorkOrderID
            Me.wiStatus = "Pending"
            Call ShowVMRS
        End If
    End Sub

Private Sub Form_Load()
   Call HideVMRS
End Sub

Private Sub HideVMRS()
Me.VMRSSearch = ""
Me.EnterMsg.Visible = False
Me.DDClick.Visible = False
Me.VMRSSearch.Visible = False
Me.VMRSList.Visible = False
Me.VMRSList.RowSource = ""
If Me.WorkOrderID & "test" <> "test" Then Me.ExpenseSum = Nz(DSum("wiLaborCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & "") + DSum("wiPartsCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & ""), 0)

End Sub

Private Sub ShowVMRS()
Me.VMRSSearch = ""
Me.EnterMsg.Visible = True
Me.DDClick.Visible = True
Me.VMRSSearch.Visible = True
Me.VMRSList.Visible = True
Me.VMRSList.RowSource = ""
Me.ExpenseSum = DSum("wiLaborCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & "") + DSum("wiPartsCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & "")
End Sub

Private Sub ItemList_Click()

  Set RST = Me.RecordsetClone
      RST.FindFirst "[WILineItemID]=" & Me.ItemList
      Me.Bookmark = RST.Bookmark
      RST.close
  Set RST = Nothing
End Sub

Private Sub SaveMe_Click()
    Dim Validate As String
    If Nz(Me.wiVMRSCode, "") = "" Then Validate = Validate & "   VMRS Code" & vbNewLine
    If Nz(Me.wiDesc, "") = "" Then Validate = Validate & "   VMRS Desc" & vbNewLine
    If Nz(Me.wiPartsCost, 0) = 0 Then Validate = Validate & "   Part Cost" & vbNewLine
    If Nz(Me.wiStatus, "") = "" Then Validate = Validate & "   Item Status" & vbNewLine
    If Validate <> "" Then
        MsgBox "You Cannot Save this Expense record." & vbNewLine & "You're Missing:" & vbNewLine & Validate & "(nothing saved)", vbCritical, "Missing Data!"
        Exit Sub
    End If
    DoCmd.RunCommand acCmdSaveRecord
    Me.ItemList.Requery
    Me.ExpenseSum = DSum("wiLaborCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & "") + DSum("wiPartsCost", "WorkorderItems", "[WorkOrderID]=" & Me.WorkOrderID & "")
End Sub
Private Sub DeleteMe_Click()

    Response = MsgBox("Are you sure you want to Delete Item:" & vbNewLine & Me.wiDesc, vbQuestion + vbYesNo + vbCritical, "Delete Item?")
    If Response = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("DELETE * From WorkorderItems Where WILineItemID=" & Me.wiLineItemID & ";")
        DoCmd.SetWarnings True
        On Error Resume Next
        DoCmd.GotoRecord , , acLast
        On Error GoTo 0
        Me.Form.Requery
        Me.ItemList.Requery
        Call HideVMRS
    End If
End Sub



Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_RmktExpense"
End Sub



Private Sub VMRSList_DblClick(Cancel As Integer)
  Me.wiVMRSCode = Me.VMRSList.Column(0)
  Me.wiDesc = Me.VMRSList.Column(1)
  Me.wiPartsCost.SetFocus
  Call HideVMRS
End Sub



Private Sub VMRSSearch_Change()
    SQLz = "Select "
    SQLz = SQLz & "VMRS9,DescrLong"
    SQLz = SQLz & " FROM vw_VMRS_Lookup"
    If IsNumeric(Left(Me.VMRSSearch.Text, 3)) Then
        SQLz = SQLz & " WHERE VMRS9 Like '" & Me.VMRSSearch.Text & "%'"
    Else
        SQLz = SQLz & " WHERE DescrLong Like '%" & Me.VMRSSearch.Text & "%'"
    End If
    Me.VMRSList.RowSource = SQLz
End Sub
