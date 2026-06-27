Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Askit, PVendor, PType, FilterX, RInvX As String
Dim SetDate As Date
Dim FilterChange As Boolean

Private Sub DeleteAllShown_Click()
    Dim SQLy As String
    Dim DWhere As String
    Dim cItemds As Long
    Dim sExpense As String
    On Error GoTo FixMe
    DWhere = Replace(Replace(Replace(SplitString(SplitString(Me.Form.RecordSource, "WHERE", 1), "Order", 0), "E.", ""), "(", ""), ")", "")
    SQLy = Replace("DELETE FROM RemarketExpense WHERE " & SplitString(SplitString(Me.Form.RecordSource, "WHERE", 1), "Order", 0), "E.", "")
    cItemds = DCount("ExpenseTypeID", "RemarketExpense", "" & DWhere)
    sExpense = DSum("Expense", "RemarketExpense", "" & DWhere)
    Askit = MsgBox("You are about to MASS DELETE Expense Records!" & vbNewLine & _
        "     Items: " & cItemds & vbNewLine & _
        "     Expense: " & Format(sExpense, "$#,##0") & vbNewLine & vbNewLine & _
        "Atlaas Core will ask again just to make sure..." & vbNewLine & _
        "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Kill it?")
    If Askit <> 6 Then Exit Sub
    MessageUser "Open", "PreParing Your DELETE Request", "You will be asked again..."
    DoCmd.SetWarnings True
    DoCmd.RunSQL (SQLy)
FixMe:
    MessageUser "Close"
    Me.Form.Requery
End Sub

Private Sub DeleteRecord_Click()
    On Error GoTo FixMe

    Askit = MsgBox("You are about to Delete an Expense Record!" & vbNewLine & _
                 "     Vendor: " & Me.VendorID.Column(1) & vbNewLine & _
                 "     Expense: " & Me.ExpenseTypeId.Column(1) & vbNewLine & _
                 "     AssetID: " & Me.AssetID & vbNewLine & _
                   "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Kill it?")
    If Askit <> 6 Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From RemarketExpense WHERE RMKTExpenseId=" & Me.RMKTExpenseId & ";")
    DoCmd.SetWarnings True
    Me.Form.Requery
FixMe:
End Sub

Private Sub ExpenseTypeId_GotFocus()
    PType = ExpenseTypeId.Column(1)
End Sub


Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.[RMKTExpenseId]
    Me.Parent.AssetID = Me.AssetID
    Me.Parent.RMKTExpenseId = Me.RMKTExpenseId
    Me.Parent.VendorID = Me.VendorID
    On Error GoTo 0
End Sub

Public Sub GetCount(Optional GetCost As Boolean)
    Dim rs As Recordset
    Dim ExpenseSum As Currency
    Set rs = Me.Recordset.Clone
    rs.MoveFirst
    Me.Parent.EventsShown = rs.RecordCount
    FilterChange = False
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    If Not IsNull(Me.Filter) Then FilterX = Me.Filter Else FilterX = ""
    If FilterX <> Me.Filter Then FilterChange = True Else FilterChange = False
End Sub

Private Sub Form_Load()
    FilterChange = False
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
    Askit = MsgBox("You are about to Change The ReimburseStatus Type." & vbNewLine & _
                 "     From Type: " & PType & vbNewLine & _
                 "     To Type: " & ReimburseStatusID.Column(1) & vbNewLine & _
                   "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change ReimburseStatus, Living on the edge.")
    If Askit <> 6 Then
        Me.Undo
    End If
End Sub

Private Sub ReimburseStatusID_GotFocus()
PType = ReimburseStatusID.Column(1)
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.AssetID)
    Me.UnitVIN.SetFocus
End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
    Call UnitID_DblClick(1)
End Sub

Private Sub ExpenseTypeId_BeforeUpdate(Cancel As Integer)

    Askit = MsgBox("You are about to Change The Expense Type." & vbNewLine & _
                 "     From Type: " & PType & vbNewLine & _
                 "     To Type: " & ExpenseTypeId.Column(1) & vbNewLine & _
                   "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change The ExpType, Are you Crazy?")
    If Askit <> 6 Then
        Me.Undo
        Exit Sub
    End If
End Sub

Private Sub VendorID_BeforeUpdate(Cancel As Integer)
    Askit = MsgBox("You are about to Change a Vendor" & vbNewLine & _
                 "     From Vendor: " & PVendor & vbNewLine & _
                 "     To Vendor: " & Me.VendorID.Column(1) & vbNewLine & _
                   "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change The Vendor, Are you Crazy?")
    If Askit <> 6 Then
        Me.Undo
        Exit Sub
    End If
End Sub

Public Sub VendorID_GotFocus()
    PVendor = Me.VendorID.Column(1)
End Sub
