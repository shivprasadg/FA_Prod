Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String
Dim Response   As String
Dim PType      As String

Private Sub AddNew_Click()
    If Me.Dirty Then
        MsgBox "You Cannot Add a record at this time." & vbNewLine & "You Must SAVE/UPDATE This One First", vbCritical, "Editing Data!"
        Exit Sub
        End
    End If
        Response = MsgBox("You are about to Add a Work Order." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Add Work Order?")
        If Response = vbYes Then
            DoCmd.GotoRecord , , acNewRec
            Me.wAssetId = Forms!Unitsform.Form.UnitID
            Me.wStartDate = Date
            Me.wStatus = "Pending"
            Me.wClientID = DLookup("ClientID", "vw_SixKeys", "AssetID=" & Me.wAssetId & "")
            Me.wClientIDGroupID = DLookup("ClientGroupID", "vw_SixKeys", "AssetID=" & Me.wAssetId & "")
            DoCmd.RunCommand acCmdSaveRecord
            Me.wNotes = "Work Order WO-" & Me.wAssetId & "-" & Me.WorkOrderID & " creadted by " & FaUserName
            Me.wONumber = "WO-" & Forms!Unitsform.Form.UnitID & "-" & Me.WorkOrderID
            Me.WorkOrderItemsSub.Visible = True
        End If
    End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

   ' If Me.Dirty Then
   ' If Nz(Me.VendorRef, 0) = 0 Or Nz(Me.ExpenseType, 0) = 0 Or Nz(Me.ExpenseDate, 0) = 0 Or Nz(Me.Expense, 0) = 0 Then
   '     Me.Undo
        'DoCmd.SetWarnings False
        'DoCmd.RunSQL ("DELETE * From RemarketExpense Where RMKTExpenseId=" & Me.RMKTExpenseId & ";")
        'DoCmd.SetWarnings True
   ' End If
  '  End If
End Sub

Private Sub Form_Load()
 If Nz(DCount("WorkOrderID", "WorkOrders", "wAssetID=" & Nz(Me.wAssetId, 0) & ""), 0) = 0 Then
     Me.AddNew.Visible = True
     'Me.WorkOrderItemsSub.Visible = False
     Me.Detail.Visible = False
     Else
     Me.AddNew.Visible = False
     Me.Detail.Visible = True
     'Me.WorkOrderItemsSub.Visible = True
 End If
End Sub

Private Sub ReimburseStatusID_AfterUpdate()
    If Me.ReimburseStatusID = 1 And Not Me.ReimburseDate Is Null Then Me.ReimburseDate = Null
    If Me.ReimburseStatusID > 2 Then
AskDate:
        SetDate = InputBox("Provide the DATE for this Event." & vbNewLine & _
                         "          Please enter that date, below." & vbNewLine & _
                         "          Today's date is the default." & vbNewLine & vbNewLine & _
                         "          ( You MUST use a date )" _
                       , "Event Date:", Date)
        If Not IsDate(SetDate) Then GoTo AskDate
        Me.ReimburseDate = SetDate
        If Me.RInvoiceNum Is Null Then
            If Me.ReimburseStatusID = 2 Or Me.ReimburseStatusID = 4 Or Me.ReimburseStatusID = 5 Then
AskRInv:
                RInvX = InputBox("Provide the Request Invoice for this Event." & vbNewLine & _
                               "          Please enter that Inv Number, below." & vbNewLine & vbNewLine & _
                               "          ( You MUST Enter a Value )" _
                      , "Request Invoice Num:", "")
                If RInvX = "" Then GoTo AskRInv
            End If
        End If
    End If
End Sub

Private Sub ReimburseStatusID_BeforeUpdate(Cancel As Integer)
    If Me.ReimburseStatusID = 4 And FAUserAccessLevel < 99 Then
    Call MsgBox("You cannot Change an already Reimbursed Expense", vbCancel + vbCritical, "No Cookies for you")
        Me.Undo
        Exit Sub
    End If
    Askit = MsgBox("You are about to Change Reimburse Status." & vbNewLine & _
                 "     From: " & PType & vbNewLine & _
                 "     To: " & ReimburseStatusID.Column(1) & vbNewLine & _
                   "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change ReimburseStatus, Living on the edge.")
    If Askit <> 6 Then
        Me.Undo
    End If
End Sub

Private Sub ReimburseStatusID_GotFocus()
PType = ReimburseStatusID.Column(1)
End Sub

Private Sub SaveMe_Click()
    Dim Validate As String
    If Nz(Me.VendorREF, 0) = 0 Then Validate = Validate & "   Vendor" & vbNewLine
    If Nz(Me.ExpenseType, 0) = 0 Then Validate = Validate & "   Expense Type" & vbNewLine
    If Nz(Me.ExpenseDate, 0) = 0 Then Validate = Validate & "   Expense Date" & vbNewLine
    If Nz(Me.Expense, 0) = 0 Then Validate = Validate & "   Expense Amount" & vbNewLine
    If Validate <> "" Then
        MsgBox "You Cannot Save this Expense record." & vbNewLine & "You're Missing:" & vbNewLine & Validate & "(nothing saved)", vbCritical, "Missing Data!"
        Exit Sub
    End If
    Me.AssetID = Forms!Unitsform.Form.UnitID
    DoCmd.RunCommand acCmdSaveRecord
    Call LoadData(Forms!Unitsform.Form.UnitID, True)

End Sub
Private Sub DeleteMe_Click()

    Response = MsgBox("Are you sure you want to Delete This Expense?", vbQuestion + vbYesNo + vbCritical, "Delete Expense?")
    If Response = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("DELETE * From RemarketExpense Where RMKTExpenseId=" & Me.RMKTExpenseId & ";")
        DoCmd.SetWarnings True
        Call LoadData(Forms!Unitsform.Form.UnitID, True)
    End If
End Sub

Private Sub ExpenseList_Click()
    SQLz = "SELECT * From RemarketExpense Where RMKTExpenseId=" & Me.ExpenseList.Column(0)
    Me.Form.RecordSource = SQLz
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_RmktExpense"
End Sub

Public Function LoadData(UnitIDx As Long, Optional ForceX As Boolean)
    Me.EnterMsg.Visible = False
    On Error Resume Next
    If Nz(Me.AssetID, 0) = 0 Or Nz(Me.AssetID, 0) <> UnitIDx Or ForceX = True Then
        On Error GoTo 0
        SQLz = "SELECT E.RMKTExpenseId, V.VendorSNm AS Vendor, T.RemarketExpense AS Type "
        SQLz = SQLz & ", CCur(E.Expense) AS Expense ,L.ReimburseStatus As RStatus, Format(ReimburseDate,""mm/dd/yy"") as RDate "
        SQLz = SQLz & "FROM (((RemarketExpense As E "
        SQLz = SQLz & "INNER JOIN RemarketExpenseType as T ON E.ExpenseTypeId = T.RemarketExpenseTypeId) "
        SQLz = SQLz & "INNER JOIN LookUp_ReimburseStatus as L ON L.ReimburseStatusID = E.ReimburseStatusID) "
        SQLz = SQLz & "LEFT JOIN Vendors As V ON E.VendorId = V.VendorID) "
        SQLz = SQLz & "WHERE AssetId=" & UnitIDx & " "
        SQLz = SQLz & "ORDER BY T.RemarketExpense;"
        Me.ExpenseList.RowSource = SQLz
        Me.Form.RecordSource = "SELECT * From RemarketExpense WHERE AssetId=" & UnitIDx & ";"
        Me.ExpenseSum = DSum("Expense", "RemarketExpense", "[AssetID]=" & UnitIDx & "")
    End If
    On Error GoTo 0
    Exit Function
FixMe:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Sub Form RmktExpense / Unit LoadData Sub", , "ATLAAS Core"
End Function
