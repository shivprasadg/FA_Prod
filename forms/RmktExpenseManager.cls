Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx, SQLz, SQLy, Askit As String
Dim SQLv As String
Dim SQLend As String
Dim NeedSearch As Integer

Private Sub ClearFilters_Click()
    Call Form_Load
    Me.EventsShown = Null
    Me.MassUpdateReimbursed.Enabled = False

End Sub

Private Sub ClearFilters2_Click()
    Call ClearFilters_Click
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DoSearch_Click()
    Call SetSQLx
    Call SetWhere
    NeedSearch = 0
    Me.RMKTExpenseSub.Form.RecordSource = SQLx & SQLz
    Call Me.RMKTExpenseSub.Form.GetCount(1)
    Me.MassUpdateReimbursed.Enabled = True
End Sub

Private Sub ExportToExcel_Click()
    If NeedSearch = 1 Then Call DoSearch_Click
    Dim XLS, TitleX As String
    Call SetSQLx
    Call SetWhere
    XLS = "|A,TL|B,cLabel:End=B:3:Vendor Inv Num|C,BC|D,$E|G,TC|G,cLabel:End=G:3:VInv Date|H,TC|H,cLabel:End=H:3:VInv Paid|I,$E|I,cLabel:End=I:3:VInv Amt|K,TX|K,BC|N,BC|P,BC|T,$E|"
    MessageUser "Open", "Exporting RMKT Expenses", "...Getting Data..."
    SendReport2Excel SQLx & SQLz, "RMKTExpenses", "RMKT Expense Export: " & Date, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub Form_Load()
    Me.AssetSearch = ""
    Me.BankSearch = ""
    Me.ClientSearch = ""
    Me.ExpenseTypeSearch = ""
    Me.InvNumSearch = ""
    Me.startDate = ""
    Me.Enddate = ""
    Me.VINSearch = ""
    Me.VendorSearch = ""
    Me.RStatusSearch = ""
    Me.UnitStatusSearch = ""
    Me.RInvoiceNumSearch = ""
    Me.RDateSearch = ""
    Me.SumExpense = Null
    Me.DateToUse = Date
    Call SetSQLx
    Me.RMKTExpenseSub.Form.RecordSource = SQLx
    NeedSearch = 0
End Sub

Public Sub SetSQLx()
    Dim xDate As String
    xDate = "12/31/" & (DatePart("yyyy", Date) - 3)
    SQLx = ""
    SQLx = SQLx & "SELECT Vendors.VendorSNm AS Vendor, E.InvoiceNumber, E.ExpenseDate, E.Expense"
    SQLx = SQLx & ", T.RemarketExpense as ExpenseType, E.Comment, E.InvoiceDate, E.InvoicePaidDate"
    SQLx = SQLx & ", E.InvoiceAmount, Clients.ClientShNm AS Client, E.AssetID, U.UnitVIN AS VIN"
    SQLx = SQLx & ", U.UnitUnitNum AS UnitNum, U.UnitStatus"
    SQLx = SQLx & ", [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitModel] & ' ' & [UnitSubType] & ' ' & [UnitType] AS UnitDesc"
    SQLx = SQLx & ", G.Assignee_Bank AS Assignee, RS.ReimburseStatus, E.ReimburseDate"
    SQLx = SQLx & ", E.RInvoiceNum, BuyOutCost, BuyOutDate "
    SQLx = SQLx & ", E.RMKTExpenseId, E.ExpenseTypeId, G.AssigneeBankID, E.ReimburseStatusID, E.VendorId "
    SQLx = SQLx & " FROM ((((((RemarketExpense As E "
    SQLx = SQLx & " INNER JOIN Units As U ON E.AssetId = U.UnitID)"
    SQLx = SQLx & " INNER JOIN Clients ON U.ClientID = Clients.clientID)"
    SQLx = SQLx & " INNER JOIN SchGrp as G ON U.GroupREF = G.SGrpID)"
    SQLx = SQLx & " INNER JOIN LookUp_ReimburseStatus As RS ON E.ReimburseStatusID = RS.ReimburseStatusID)"
    SQLx = SQLx & " INNER JOIN RemarketExpenseType as T ON E.ExpenseTypeId = T.RemarketExpenseTypeID)"
    SQLx = SQLx & " INNER JOIN RMKT As R ON R.UnitRef = U.UnitID)"
    SQLx = SQLx & " LEFT JOIN Vendors ON E.VendorID = Vendors.VendorID "
    SQLx = SQLx & " WHERE (((E.ExpenseDate)>#" & CDate(xDate) & "#)) "

End Sub
Public Sub SetWhere()
    SQLz = ""
    If Me.AssetSearch & "test" <> "test" Then SQLz = SQLz & "AND E.AssetID=" & Me.AssetSearch & " "
    If Me.BankSearch & "test" <> "test" Then SQLz = SQLz & "AND G.AssigneeBankID=" & Me.BankSearch & " "
    If Me.ClientSearch & "test" <> "test" Then SQLz = SQLz & "AND U.ClientID=" & Me.ClientSearch & " "
    If Me.ExpenseTypeSearch & "test" <> "test" Then SQLz = SQLz & "AND E.ExpenseTypeID = " & Me.ExpenseTypeSearch & " "
    If Me.InvNumSearch & "test" <> "test" Then SQLz = SQLz & "AND E.InvoiceNumber='" & Me.InvNumSearch & "' "
    If Me.RStatusSearch.Column(0) & "test" <> "test" Then SQLz = SQLz & "AND E.ReimburseStatusID=" & Me.RStatusSearch & " "
    If Me.UnitStatusSearch & "test" <> "test" Then SQLz = SQLz & "AND U.UnitStatus='" & Me.UnitStatusSearch & "' "
    If Me.RInvoiceNumSearch & "test" <> "test" Then SQLz = SQLz & "AND E.RInvoiceNum='" & Me.RInvoiceNumSearch & "' "
    If Me.RDateSearch & "test" <> "test" Then SQLz = SQLz & "AND ReimburseDate=#" & Me.RDateSearch & "# "
    If Me.VendorSearch.Column(0) & "test" <> "test" Then SQLz = SQLz & "AND E.VendorID=" & Me.VendorSearch & " "
    If Me.VINSearch & "test" <> "test" Then SQLz = SQLz & "AND U.UnitVIN Like'*" & Me.VINSearch & "*' "

    If IsDate(Me.startDate) And IsDate(Me.Enddate) Then
        SQLz = SQLz & "AND  ExpenseDate >=#" & Me.startDate & "# AND  ExpenseDate <=#" & Me.Enddate & "# "
        Header = Header & " and By Date Range: " & Me.startDate & " ~ " & Me.Enddate
    End If
    If IsDate(Me.startDate) And Not IsDate(Me.Enddate) Then
        SQLz = SQLz & "AND  InvoiceDate =#" & Me.startDate & "# "
        Header = Header & " and Invoice Date: " & Me.startDate
    End If
    If Not IsDate(Me.startDate) And IsDate(Me.Enddate) Then
        SQLz = SQLz & " AND ExpenseDate =#" & Me.Enddate & "# "
        Header = Header & " and ExpenseDate: " & Me.Enddate
    End If
    If SQLz = "" Then NeedSearch = 0 Else NeedSearch = 1
    SQLz = SQLz & " ORDER BY InvoiceNumber, E.AssetID, ExpenseDate, ClientShNm"
End Sub


Public Sub MassUpdateReimbursed_Click()
Dim msg As String
    If Me.RStatusToUse & "test" = "test" Then
        MsgBox "You MUST pick a Reimburse Status value" & vbNewLine & "No updates were made.", vbCritical, "No Nulls"
        Exit Sub
    End If
    If Not IsDate(Me.DateToUse) Then Me.DateToUse = Date
    If RStatusToUse = 2 And Me.RInvoiceNumToUse & "test" = "test" Then
        MsgBox "To Send in a Request, you MUST Set Request Invoice Number to Use " & vbNewLine & "No updates were made.", vbCritical, "No Invoice, No Service"
        Exit Sub
    End If
    msg = "You are about to update " & Me.EventsShown & " Reimbursibe Expenses."
    msg = msg & vbNewLine & "with Reimburse Status: '" & Me.RStatusToUse.Column(1)
    msg = msg & vbNewLine & "Using Date: " & Me.DateToUse
    If Me.RInvoiceNumToUse & "test" <> "test" Then msg = msg & vbNewLine & "with Inv Num: " & Me.RInvoiceNumToUse
    If FAUserAccessGroup Like "*ACCT*" Then msg = msg & vbNewLine & "only SB units will change"
    msg = msg & vbNewLine & vbNewLine & "          Do you agree?"
    msg = msg & vbNewLine & vbNewLine & "Note: Once an expense is 'Aprroved' it cannot change."
    msg = msg & vbNewLine & "           Contact the IT Dept, if need be."
    Askit = MsgBox(msg, vbQuestion + vbYesNo + vbCritical, "Mass Update")
    If Askit <> 6 Then GoTo FixMe:
    Call SetSQLx
    'Call SetWhere
    SQLv = SQLx
    SQLv = SplitString(SQLv, "FROM", 1)
    SQLv = "Select RMKTExpenseId From RemarketExpense Where RMKTExpenseId IN (SELECT E.RMKTExpenseId FROM " & SQLv & SQLz
    SQLv = SplitString(SQLv, "ORDER BY ", 0)
    If FAUserAccessGroup Like "*ACCT*" And FAUserAccessLevel < 99 Then SQLv = SQLv & " And UnitStatus = 'SB' " Else SQLv = SQLv & " And UnitStatus <> 'SB' "
    SQLv = SQLv & ")"
    SQLy = "Update RemarketExpense Set ReimburseStatusID = " & Me.RStatusToUse & ""
    If Me.RStatusToUse = 1 Then
        SQLy = SQLy & ", ReimburseDate = Null "
    Else
        SQLy = SQLy & ", ReimburseDate = '" & Me.DateToUse & "' "
    End If
    If Me.RInvoiceNumToUse & "test" <> "test" Then SQLy = SQLy & ", RInvoiceNum = '" & Me.RInvoiceNumToUse & "' "
    SQLy = SQLy & "WHERE RMKTExpenseId IN (" & SQLv & ") "
    If FAUserAccessLevel < 99 Then SQLy = SQLy & " And ReimburseStatus <> 4 "
    On Error GoTo FixMe
    DoCmd.RunSQL (SQLy)
    NeedSearch = 0
    Me.RMKTExpenseSub.Form.RecordSource = SQLx & SQLz
    Exit Sub
FixMe:
    MsgBox "Update to Reimbursed was Canceled.", vbInformation, "Wisdom is Good"
    Me.RMKTExpenseSub.Form.RecordSource = SQLx & SQLz
End Sub

Private Sub OpenBuyer_Click()
    DoCmd.OpenForm "VendorsForm", , , "Vendorid=" & Me.VendorID
End Sub

Private Sub OpenSchedule_Click()
    If IsNull(Me.AssetID) Then Exit Sub
    DoCmd.OpenForm "ScheduleControl"
    Forms!ScheduleControl.Form.SearchAssetID2.SetFocus
    Forms!ScheduleControl.Form.SearchAssetID2.Text = Me.AssetID
    Forms!ScheduleControl.Form.UnitLookUp = Me.AssetID
    Form_ScheduleControl.Remotelookup

End Sub
Private Sub OpenUnit_Click()
    If IsNull(Me.AssetID) Then Exit Sub
    Call OpenUnitsForm(Me.AssetID)
End Sub
Private Sub ClientSearch_Change()
    NeedSearch = 1
End Sub
Private Sub AssetSearch_Change()
    NeedSearch = 1
End Sub
Private Sub BankSearch_Change()
    NeedSearch = 1
End Sub
Private Sub ExpenseTypeSearch_Change()
    NeedSearch = 1
End Sub
Private Sub InvNumSearch_Change()
    NeedSearch = 1
End Sub
Private Sub EndDate_Change()
    NeedSearch = 1
End Sub
Private Sub ReSearch_Change()
    NeedSearch = 1
End Sub
Private Sub StartDate_Change()
    NeedSearch = 1
End Sub
Private Sub VendorSearch_Change()
    NeedSearch = 1
End Sub
Private Sub VINSearch_Change()
    NeedSearch = 1
End Sub
