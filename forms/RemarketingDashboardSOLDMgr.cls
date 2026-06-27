Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim XLS, SqlA, SQLx, strSQL, TargetX, SourceCompanyX, YearX, UnitTypeX, MakeX, ModelX, BuyerX, StatusX, OrderByx, SoldDateRangeX, BuyOutRangeX, InvoiceSentRangeX, CityX, StateX As String
Dim SQLz As String
Dim intRow, i, Countx, UnitCountx As Integer
Dim RST As DAO.Recordset

Private Sub ExportToExcel_Click()
    Dim XLS, Sqlr, TitleX, UnitIDx As String
    Dim i As Integer
    TitleX = "Admin Report " & Date
    Sqlr = "SELECT SourceCompany AS [Account Name]"
    Sqlr = Sqlr & ", Schedule AS [Lease Sch]"
    Sqlr = Sqlr & ", UnitGroup AS [Group]"
    Sqlr = Sqlr & ", UnitID"
    Sqlr = Sqlr & ", UnitNum"
    Sqlr = Sqlr & ", UnitVIN"
    Sqlr = Sqlr & ", UnitStatus"
    Sqlr = Sqlr & ", RStatus AS InvStatus"
    Sqlr = Sqlr & ", UnitYr AS [Year]"
    Sqlr = Sqlr & ", UnitMake AS Make"
    Sqlr = Sqlr & ", UnitModel AS Model"
    Sqlr = Sqlr & ", UnitType AS Type"
    Sqlr = Sqlr & ", Odometer"
    Sqlr = Sqlr & ", PCity"
    Sqlr = Sqlr & ", PState"
    Sqlr = Sqlr & ", Buyer"
    Sqlr = Sqlr & ", Disposition AS Dispn"
    Sqlr = Sqlr & ", SalesPerson"
    Sqlr = Sqlr & ", TotalExpense as RMKT_Exp"
    Sqlr = Sqlr & ", Residual"
    Sqlr = Sqlr & ", ResidualPct"
    Sqlr = Sqlr & ", BookValue"
    Sqlr = Sqlr & ", TargetSalePrice"
    Sqlr = Sqlr & ", UnitSalePrice AS [Sale Price]"
    Sqlr = Sqlr & ", UnitSalePricePct AS [Sale Pct]"
    Sqlr = Sqlr & ", Buyout AS [BuyOut Cost]"
    Sqlr = Sqlr & ", BuyOutPct"
    Sqlr = Sqlr & ", TermPenalty AS [Term Penalty]"
    Sqlr = Sqlr & ", TotalProfit"
    Sqlr = Sqlr & ", TotalExpense AS RMKTExpense"
    Sqlr = Sqlr & ", GainLoss"
    Sqlr = Sqlr & ", ParkDate"
    Sqlr = Sqlr & ", DaysInInventory AS [Days In]"
    Sqlr = Sqlr & ", [SUF Returned]"
    Sqlr = Sqlr & ", LeaseExp"
    Sqlr = Sqlr & ", OffLeaseDate"
    Sqlr = Sqlr & ", BuyOutDate"
    Sqlr = Sqlr & ", DaysInAcct AS [Days BuyOut]"
    Sqlr = Sqlr & ", InvoiceSent"
    Sqlr = Sqlr & ", InvoicePaid"
    Sqlr = Sqlr & ", TitleRecv"
    Sqlr = Sqlr & ", BankPartner"
    Sqlr = Sqlr & ", BankShareNet"
    Sqlr = Sqlr & ", AgeAtSale"
    Sqlr = Sqlr & ", ClientID"
    Sqlr = Sqlr & ", LocationID"
    Sqlr = Sqlr & ", BuyerID"
    Sqlr = Sqlr & " FROM RemarketingReportAllUnits "
    Sqlr = Sqlr & " Where UnitID in ("

    Dim RST As Recordset, intI As Integer
    Dim Rec As Record
    Set RST = Me.RecordsetClone
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        UnitIDx = UnitIDx & RST!UnitID & ","
        RST.MoveNext
    Next i
    UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
    Sqlr = Sqlr & UnitIDx & ")"
    ' Debug.Print UnitIDx
    XLS = "|A,TL|D,TX|H,TC|G,BC|M,Sint|R,BC|S,$E|T,$E|U,P2|V,$E|W,$E|X,$E|Y,P2|Z,$E|AA,P2|AB,$E|AC,$E|AD,$E|AE,$E|AG,BD|AQ,$E|"
    MessageUser "Open", "Getting Data...", "This will be quick"
    SendReport2Excel Sqlr, "Admin Report", TitleX, True, XLS, 0, False, , , , "RMKTInventoryReport"
    MessageUser "Close"
End Sub

Private Sub Form_Load()
    Me.Form.FilterOn = False
    Call SetYearList
    Call ClearBoxes
    Call SetButtonYears
    Call cmdSetFilters_Click
End Sub

Private Sub Odometer_Click()
Me.ODOMETER.ControlTipText = "ODO mileage collected from " & DLookup("ODOSource", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID) & " on " & DLookup("ODOReadDate", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID)
End Sub

Private Sub ReportsList_Click()
    Me.ReportsList.ControlTipText = Nz(ReportsList.Column(1), "test")
End Sub

Private Sub ReportsList_DblClick(Cancel As Integer)
If Coalesce(Me.UnitsShown, 0) < 1 Then
MsgBox "You're Tring To Run A Report With NO Units Showing." & vbNewLine & "Please Change The Filters or Reset the Form", vbCritical, "No Shirts No Shoes, No Service"
Exit Sub
End If
    Dim FormX, CommandX As String
    FormX = Me.ReportsList.Column(3)
    CommandX = Me.ReportsList.Column(2)
    If FormX = "Local" Then
        CallByName Me, CommandX, VbMethod
    Else
        CallByName FormX, CommandX, VbMethod
    End If
End Sub
Public Sub JFReports()
    Dim XLS, FormulaX As String
    Dim UnitCountx As Integer
    Dim Showx, Lastx As Boolean
    Dim Rptx As Integer

    MessageUser "Open", "Creating Excel Inventory Report", "...Getting Data..."
    xx = 1
    Showx = False
    Rptx = 1
    Lastx = False
    Call cmdSetFilters_Click
Inv_Added_As_Of_Report:
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND BuyOutDate Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# AND UnitStatus not like '*T*' Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo 0
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating Inventory As Of Date Report", "...This will take just a moment..."
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,$R|AA,ZZ|AA,ZZ|AA,ZZ|AB,ZZ|AB,ZZ|"

    If UnitCountx > 0 Then
        SendReport2Excel SqlA, "Inv_Added_Date", "Remarketing Inv Added During " & SoldDateStart & "~" & SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, Showx, Rptx, Lastx
        Rptx = Rptx + 1
    End If

Inventory_At_Date:
    MessageUser "Update", "Creating Excel Inventory Report", "...Getting Data..."
    Call cmdSetFilters_Click
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND (isnull(InvoicePaid) or (InvoicePaid > #" & Me.SoldDateEnd & "# AND InvoicePaid < #" & Me.SoldDateStart & "#)) AND UnitStatus not like '*T*' Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo 0
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating Inventory @ Date Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"
    If UnitCountx > 0 Then
        SendReport2Excel SqlA, "Inventory_@_Date", "Remarketing Inventory at " & Me.SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, Showx, Rptx, Lastx
        Rptx = Rptx + 1
    End If
Inventory_TR:
    MessageUser "Update", "Creating TR Units Report", "...Getting Data..."
    Call cmdSetFilters_Click
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null  AND [UnitStatus] = 'TR'  Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo 0
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating TR Inventory Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"

    If UnitCountx > 0 Then
        SendReport2Excel SqlA, "TR_Units", "Remarketing TR Units | " & Date & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, Showx, Rptx, Lastx
        Rptx = Rptx + 1
    End If

Inventory_Current:
    MessageUser "Update", "Creating Excel Inventory Report", "...Getting Data..."
    Call cmdSetFilters_Click
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null  AND ([UnitStatus] = 'R' OR [UnitStatus] = 'PS') Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo 0
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Getting Current Inventory Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"
    If UnitCountx > 0 Then
        SendReport2Excel SqlA, "Current_Inventory", "Remarketing Inventory at " & Date & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, Showx, Rptx, Lastx
        Rptx = Rptx + 1
    End If
    Lastx = "True"
ExportSoldReport:
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportSoldUnits")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND InvoicePaid Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo 0
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating Excel Sold Report", "...This will take just a moment..."
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|M,TC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|X,BC|AA,ZZ|AB,ZZ|AB,ZZ|"
    SendReport2Excel SqlA, "Sold_Units", "Remarketing Sold Units " & SoldDateStart & "~" & SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, Showx, Rptx, Lastx
    MessageUser "Close"

End Sub
Private Sub SetYearList()
    Dim YearListX As String
    yearlist = Chr(34) & "ALL" & Chr(34) & ";"">0" & Chr(34)
    Dim x As Integer
    For x = CInt(DatePart("yyyy", Date) + 1) To CInt(DatePart("yyyy", Date) - 17) Step -1
        yearlist = yearlist & ";" & Chr(34) & x & Chr(34) & ";" & Chr(34) & "=" & x & Chr(34)
    Next x
    Me.lstModelYear.RowSource = yearlist
End Sub
Public Sub TargetPriceSetBuild()
    TargetX = "XX"
    If Me.TargetValueList = 1 Then TargetX = " AND (TargetSalePrice > 0)"
    If Me.TargetValueList = 2 Then TargetX = " AND (TargetSalePrice = 0)"
    If Me.TargetValueList = 3 Then TargetX = " AND (FMV > 0)"
    If Me.TargetValueList = 4 Then TargetX = " AND (FMV = 0)"
    If Me.TargetValueList = 5 Then TargetX = " AND (FMVBase > 0)"
    If Me.TargetValueList = 6 Then TargetX = " AND (FMVBase = 0)"
    If Me.TargetValueList = 7 Then TargetX = " AND (TermPenalty <> 0)"
    If Me.TargetValueList = 8 Then TargetX = " AND (TermPenalty = 0)"
End Sub
Private Sub ResetForm_Click()
' Call SetDefault
'Call SETSQL
    Call ClearBoxes
    Call ClearAll_Click
    Call cmdSetFilters_Click
End Sub
Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub

Private Sub ClearAll_Click()
    Call ClearList("lstSourceCompany")
    Call ClearList("lstModelYear")
    Call ClearList("lstUnitType")
    Call ClearList("lstMake")
    Call ClearList("lstModel")
    Call ClearList("lstBuyer")
    Call ClearList("lstStatus")
    Call ClearList("TargetValueList")
  '  Call ClearList("CityList")
   ' Call ClearList("StateList")
    ClearDates (1)
    ClearDates (2)
    ClearDates (3)
End Sub
Private Sub ClearAndRun_Click()
    Call ClearAll_Click
    Call cmdSetFilters_Click
End Sub

Private Sub ClearClients_Click()
    Call ClearList("lstSourceCompany")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearYear_Click()
    Call ClearList("lstModelYear")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearUnitType_Click()
    Call ClearList("lstUnitType")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearMake_Click()
    Call ClearList("lstMake")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearModel_Click()
    Call ClearList("lstModel")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearBuyer_Click()
    Call ClearList("lstBuyer")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearStatus_Click()
    Call ClearList("lstStatus")
    Call cmdSetFilters_Click
End Sub

'Private Sub ClearCity_Click()
'    Call ClearList("CityList")
'    Call cmdSetFilters_Click
'End Sub
'Private Sub ClearState_Click()
'    Call ClearList("StateList")
'    Call cmdSetFilters_Click
' End Sub

Private Sub ClearTarget_Click()
    Call ClearList("TargetValueList")
    Call TargetPriceSetBuild
    ' Call SETSQL
End Sub
Public Sub ClearList(ListBoxName As String)
    Dim ItemsX As Variant
    With Me(ListBoxName)
        For Each ItemsX In .ItemsSelected
            .Selected(ItemsX) = False
        Next ItemsX
        Me(ListBoxName) = Null
    End With
End Sub

Public Sub SetDefault()
    SqlA = "SELECT * FROM RemarketingReportAllUnits WHERE (([UNITId] is not null) "
    SourceCompanyX = "AND ClientId is not null "
    YearX = "AND UnitYr is not null "
    UnitTypeX = "AND UnitType is not null "
    MakeX = "AND UnitMake is not null "
    ModelX = "AND UnitModel is not null "
    BuyerX = "AND BuyerId is not null "
    StatusX = "AND UnitStatus is not null "
    SoldDateRangeX = "XX"
    BuyOutRangeX = "XX"
    InvoiceSentRangeX = "XX"
    OrderByx = "ORDER BY SourceCompany"
    CityX = "XX"
    StateX = "XX"
End Sub

Public Sub SETSQL()
    SqlA = "SELECT * FROM RemarketingReportAllUnits WHERE [UnitID] is not null "
    SQLz = SqlA & SourceCompanyX & YearX & UnitTypeX & MakeX & ModelX & BuyerX & StatusX & " "
    If TargetX <> "XX" Then Sqlr = Sqlr & TargetX
    If SoldDateRangeX <> "XX" Then Sqlr = Sqlr & SoldDateRangeX
    If BuyOutRangeX <> "XX" Then Sqlr = Sqlr & BuyOutRangeX
    'If InvoiceSentRangeX <> "XX" Then SQLr = SQLr & InvoiceSentRangeX
    If CityX <> "XX" Then Sqlr = Sqlr & CityX
    If StateX <> "XX" Then Sqlr = Sqlr & StateX
    If BuyerX = "" Then OrderByx = "Order By SourceCompany" Else OrderByx = "ORDER BY Buyer"
    Sqlr = Sqlr & OrderByx & ";"
    Dim WhereX As String
    WhereX = Split(SQLz, "RemarketingReportAllUnits ")(1)
    WhereX = Split(WhereX, "Order")(0)

   ' Call GetSelectedCount("CityList")
   ' Me.CityList.RowSource = "SELECT PCity From RemarketingReportAllUnits " & WhereX & " GROUP BY PCity ORDER BY PCity;"
   ' If Countx > 0 Then Call ReSelectItems("CityList")

    'Call GetSelectedCount("StateList")
    'Me.StateList.RowSource = "SELECT PState From RemarketingReportAllUnits " & WhereX & " GROUP BY PState ORDER BY PState;"
    'If Countx > 0 Then Call ReSelectItems("StateList")

    Call GetSelectedCount("lstSourceCompany")
    Me.lstSourceCompany.RowSource = "SELECT SourceCompany, ClientId From RemarketingReportAllUnits " & WhereX & " GROUP BY SourceCompany, ClientId ORDER BY SourceCompany;"
    If Countx > 0 Then Call ReSelectItems("lstSourceCompany")

    Call GetSelectedCount("lstUnitType")
    Me.lstUnitType.RowSource = "SELECT UnitType From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitType ORDER BY UnitType;"
    If Countx > 0 Then Call ReSelectItems("lstUnitType")

    Call GetSelectedCount("lstMake")
    Me.lstMake.RowSource = "SELECT UnitMake From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitMake ORDER BY UnitMake;"
    If Countx > 0 Then Call ReSelectItems("lstMake")

    Call GetSelectedCount("lstModel")
    Me.LstModel.RowSource = "SELECT UnitModel From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitModel ORDER BY UnitModel;"
    If Countx > 0 Then Call ReSelectItems("lstModel")

    Call GetSelectedCount("lstModelYear")
    Me.lstModelYear.RowSource = "SELECT UnitYr From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitYr ORDER BY UnitYr desc;"
    If Countx > 0 Then Call ReSelectItems("lstModelYear")

    Call GetSelectedCount("lstBuyer")
    Me.lstBuyer.RowSource = "SELECT BuyerName,BuyerId From RemarketingReportAllUnits " & WhereX & " GROUP BY BuyerName, BuyerId ORDER BY BuyerName;"
    If Countx > 0 Then Call ReSelectItems("lstBuyer")
    'Me.lstStatus.RowSource = "SELECT UnitStatus From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitStatus ORDER BY UnitStatus;"
    Me.Form.RecordSource = SQLz & Sqlr
   ' Me.UnitCount.Requery
    Me.ReportsList.Requery
   ' UnitCountX = Me.UnitCount.Value
End Sub

Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Group1" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub
Public Sub SetButtonYears()
    Dim x As Integer
    Dim bName As String
    For x = 0 To 2
        '  bName = "Buy" & X
        '  Me(bName).Caption = Right(DatePart("yyyy", Date) - X,2)
        '  bName = "Sent" & X
        '  Me(bName).Caption = Right(DatePart("yyyy", Date) - X,2)
        bName = "Sold" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 2)
    Next x
End Sub

Private Sub cmdSetFilters_Click()
    Call SourceCompanyBuild
    Call TargetPriceSetBuild
    Call ModelYearBuild
    Call UnitTypeBuild
    Call MakeBuild
    Call ModelBuild
    Call BuyerBuild
    Call StatusBuild
    ' Call BuyOutStart_AfterUpdate
    Call SoldDateStart_AfterUpdate
    ' Call InvoiceSentStart_AfterUpdate
    'Call CityBuild
    'Call StateBuild
    Call SETSQL
End Sub
Public Sub ReSelectItems(ListBoxName As String)
    With Me(ListBoxName)
        For intRow = 0 To .ListCount - 1
            .Selected(intRow) = True
        Next intRow
    End With
End Sub
Public Sub GetSelectedCount(ListBoxName As String)
    Countx = 0
    With Me(ListBoxName)
        For intRow = 0 To .ListCount - 1
            If .Selected(intRow) Then Countx = Countx + 1
        Next intRow
    End With
End Sub
Private Sub SourceCompanyBuild()
    On Error Resume Next
    If IsNull(lstSourceCompany.Column(1)) = False Or Nz(lstSourceCompany.Column(1), "") <> "" Then
        strSQL = " AND ("
        With Me.lstSourceCompany
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "[ClientId]=" & .Column(1, intRow) & " OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([ClientId] is not null) "
    End If
    On Error GoTo 0
    SourceCompanyX = strSQL
End Sub
Private Sub ModelYearBuild()
On Error Resume Next
    If Nz(lstModelYear.Column(0), 0) <> 0 Then
        strSQL = " AND ("
        With Me.lstModelYear
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "[UnitYr]=" & .Column(0, intRow) & " OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitYr] > 1) "
    End If
    On Error GoTo 0
    YearX = strSQL
End Sub
Private Sub UnitTypeBuild()
    If Nz(Me.lstUnitType.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstUnitType
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([UnitType]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitType] <> '') "
    End If
    UnitTypeX = strSQL
End Sub
Private Sub MakeBuild()
    If Nz(lstMake.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstMake
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([UnitMake]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitMake] <> '')"
    End If
    MakeX = strSQL
End Sub
Private Sub ModelBuild()
    If Nz(LstModel.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.LstModel
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([UnitModel]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitModel] <> '')"
    End If
    ModelX = strSQL
End Sub

Private Sub BuyerBuild()
    If Nz(lstBuyer.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstBuyer
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([BuyerId]=" & .Column(1, intRow) & ") OR "
                End If
            Next intRow
        End With

        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = ""
    End If
    BuyerX = strSQL
End Sub

Private Sub StatusBuild()
    If Nz(lstStatus.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstStatus
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([UnitStatus]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        ' strip off leading comma and space
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitStatus] Like '*S*') "
    End If
    StatusX = strSQL
End Sub
Private Sub CityBuild()
    If Nz(Me.CityList.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.CityList
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([PCity]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = "XX"
        '    strSQL = " AND ([PCity] <> '')"
    End If
    CityX = strSQL
End Sub
Private Sub StateBuild()
    If Nz(Me.StateList.Column(0), "") & "test" <> "test" Then
        strSQL = " AND ("
        With Me.StateList
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([PState]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = "XX"
        ' strSQL = " AND ([PState] <> '')"
    End If
    StateX = strSQL
End Sub
Private Sub Buyer_DblClick(Cancel As Integer)
    If Me.Buyer & "TEST" = "TEST" Then Exit Sub
    Call OpenBuyersFrom(Me.BuyerId, 0)
End Sub
Private Sub Buy0_Click()
    Call BuyOutStart_Click
    Call SetDateRange("BuyOutStart", "BuyOutEnd", 0)
End Sub
Private Sub Buy1_Click()
    Call BuyOutStart_Click
    Call SetDateRange("BuyOutStart", "BuyOutEnd", 1)
End Sub
Private Sub Buy2_Click()
    Call BuyOutStart_Click
    Call SetDateRange("BuyOutStart", "BuyOutEnd", 2)
End Sub
Private Sub Sent0_Click()
    Call InvoiceSentStart_Click
    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 0)
End Sub
Private Sub Sent1_Click()
    Call InvoiceSentStart_Click
    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 1)
End Sub
Private Sub Sent2_Click()
    Call InvoiceSentStart_Click
    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 2)
End Sub
Private Sub Sold0_Click()
    Call SoldDateStart_Click
    Call SetDateRange("SoldDateStart", "SoldDateEnd", 0)
End Sub
Private Sub Sold1_Click()
    Call SoldDateStart_Click
    Call SetDateRange("SoldDateStart", "SoldDateEnd", 1)
End Sub
Private Sub Sold2_Click()
    Call SoldDateStart_Click
    Call SetDateRange("SoldDateStart", "SoldDateEnd", 2)
End Sub
Private Sub SetDateRange(StartField As String, EndField As String, YearBack As Integer)
    Me(StartField) = "1/1/" & DatePart("yyyy", Date) - YearBack
    If YearBack = 0 Then Me(EndField) = Date Else Me(EndField) = "12/31/" & DatePart("yyyy", Date) - YearBack
    If StartField Like "*Sold*" Then Call SoldDateStart_AfterUpdate
    If StartField Like "*Sent*" Then Call InvoiceSentStart_AfterUpdate
    If StartField Like "*Buy*" Then Call BuyOutStart_AfterUpdate
End Sub
Private Sub SourceCompany_DblClick(Cancel As Integer)
    If Me.SourceCompany & "TEST" = "TEST" Then Exit Sub
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
    Me.UnitID.SetFocus
End Sub
Private Sub TargetValueList_Click()
    Call TargetPriceSetBuild
    'Call SETSQL
End Sub
Private Sub UnitID_DblClick(Cancel As Integer)
    DoCmd.close acForm, "UnitsForm"
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UnitId=" & Me.UnitID & ""
      Me.UnitID.SetFocus
End Sub
'_________________________
'XLS = [Col letter],[Format Option][Command]
'Col Letter:        Number/letter of the column
'Format Options:    T = Text | B = Bold | $ = Currency | S = Standard Numbers | P = Percent (P# = Percent and the # of decimials)
'Command:           C = Center | L=Left | R=Right | E = SubTotal Col | T= Total Col Count | X = COUNT | ZZ = Delete column | HH = Hide Column
Public Sub InventoryReportMgr()
    MessageUser "Open", "Creating Manager's INV Report", "...This will take just a moment..."
    cmdSetFilters_Click
    Call Remarketing_INVReport_Mgr(SplitString(SQLz, "WHERE", 1))
    MessageUser "Close"
End Sub
Public Sub InventoryReportSales()
    MessageUser "Open", "Creating Sales' INV Report", "...This will take just a moment..."
    cmdSetFilters_Click
    SQLz = Replace(SQLz, "RemarketingReportAllUnits", "RemarketingDashboardINV_Sales")
    SQLz = Replace(SQLz, "SourceCompany", "[Account Name]")

    Call Remarketing_INVReport_Sales(SQLz, 0)
    MessageUser "Close"
End Sub
Public Sub SoldReportMgr()
    MessageUser "Open", "Creating Manager's Sold Report", "...This will take just a moment..."
    Call cmdSetFilters_Click
    SQLz = Replace(SQLz, "RemarketingReportAllUnits", "RemarketingDashboardSold_MGR")
    SQLz = Replace(SQLz, "SourceCompany", "[Account Name]")

    Call Remarketing_SOLDReport_Mgr(SQLz, 0)
    MessageUser "Close"
End Sub
Public Sub SoldReportSales()
    MessageUser "Open", "Creating Sold Report", "...This will take just a moment..."
    Call cmdSetFilters_Click
    'SQLz = Replace(SQLz, "RemarketingReportAllUnits", "RemarketingDashboardSold_Sales")
    'SQLz = Replace(SQLz, "SourceCompany", "[Account Name]")
    Call Remarketing_SOLDReport_Sales(SplitString(SQLz, "WHERE ", 1))
    MessageUser "Close"
End Sub
Public Sub AdminReport()
    MessageUser "Open", "Creating Admin Report", "...This will take just a moment..."
    Call cmdSetFilters_Click
    SQLz = Replace(SQLz, "RemarketingReportAllUnits", "RemarketingDashboard_Admin")
    SQLz = Replace(SQLz, "SourceCompany", "[Account Name]")
    SQLz = Replace(SQLz, "UnitId", "AssetId")
    SQLz = Replace(SQLz, "UnitYr", "Year")
    SQLz = Replace(SQLz, "UnitMake", "Make")
    SQLz = Replace(SQLz, "UnitModel", "Model")
    SQLz = Replace(SQLz, "UnitType", "Type")

    Call Remarketing_AdminReport(SQLz, 0)
    MessageUser "Close"
End Sub
Public Sub AssigneeCostReport()
Exit Sub
    MessageUser "Open", "Assignee Cost", "...Getting data..."
    cmdSetFilters_Click
    MessageUser "Update", "Assignee Cost", "...Setting SQL..."
    MessageUser "Update", "Assignee Cost", "...this will take a few moments..."
    SQLz = Replace(SQLz, "RemarketingReportAllUnits", "Remarketing_Assignee_Cost_Report")
    'XLS = "|B,TC|C,TC|F,TX|F,TC|G,TC|L,$E|M,$E|O,BC|T,BL|V,$E|W,$E|X,$E|Y,$E|Z,$E|AA,HH|"
    SendReport2Excel SQLz, "Sold PnL", "Sold Units " & Format(UnitCountx, "#,#") & " Units " & Format(Date, "mm-dd-yy"), True, XLS, UnitCountx
    MessageUser "Close"
End Sub
Private Sub InvoiceSentStart_AfterUpdate()
'If IsDate(Me.InvoiceSentStart) And IsDate(Me.InvoiceSentEnd) Then
'    InvoiceSentRangeX = "And InvoiceSent Between #" & Me.InvoiceSentStart & "# AND #" & Me.InvoiceSentEnd & "# "
'    OrderByX = "ORDER BY ParkedDate"
'    Call SETSQL
'Else
'    InvoiceSentRangeX = "XX"
'    OrderByX = "ORDER BY SourceCompany"
'End If
'If Not IsDate(Me.InvoiceSentEnd) And Not IsDate(Me.InvoiceSentStart) Then Call SETSQL
End Sub
Private Sub InvoiceSentEnd_AfterUpdate()
    Call InvoiceSentStart_AfterUpdate
End Sub
Private Sub BuyOutStart_AfterUpdate()
' If IsDate(Me.BuyOutEnd) And IsDate(Me.BuyOutStart) Then
'    BuyOutRangeX = "And BuyOutDate Between #" & Me.BuyOutStart & "# AND #" & Me.BuyOutEnd & "# "
'    OrderByX = "ORDER BY BuyOutDate"
'    Call SETSQL
'Else
'     BuyOutRangeX = "XX"
'     OrderByX = "ORDER BY SourceCompany"
'End If
' If Not IsDate(Me.BuyOutEnd) And Not IsDate(Me.BuyOutStart) Then Call SETSQL
End Sub
Private Sub BuyOutEnd_AfterUpdate()
    Call BuyOutStart_AfterUpdate
End Sub
Private Sub SoldDateStart_AfterUpdate()
    If IsDate(Me.SoldDateEnd) And IsDate(Me.SoldDateStart) Then
        SoldDateRangeX = "And InvoicePaid Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# "
        OrderByx = "ORDER BY InvoicePaid"
        '  Call SETSQL
    Else
        SoldDateRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
    ' If Not IsDate(Me.SoldDateEnd) And Not IsDate(Me.SoldDateStart) Then Call SETSQL
End Sub
Private Sub SoldDateEnd_AfterUpdate()
    SoldDateStart_AfterUpdate
End Sub

Private Sub SoldDateStart_Click()
    ClearDates (1)
    ClearDates (2)
End Sub
Private Sub SoldDateEnd_Click()
    ClearDates (1)
    ClearDates (2)
End Sub
Private Sub InvoiceSentStart_Click()
    ClearDates (3)
    ClearDates (2)
End Sub
Private Sub InvoiceSentEnd_Click()
    ClearDates (3)
    ClearDates (2)
End Sub
Private Sub BuyOutStart_Click()
    ClearDates (1)
    ClearDates (3)
End Sub
Private Sub BuyOutEnd_Click()
    ClearDates (1)
    ClearDates (3)
End Sub

Private Sub ClearDates(LevelX As Integer)
    If LevelX = 1 Then
        '   Me.InvoiceSentStart = ""
        '   Me.InvoiceSentEnd = ""
        '   InvoiceSentRangeX = "XX"
    End If
    If LevelX = 2 Then
        '  Me.BuyOutStart = ""
        '  Me.BuyOutEnd = ""
        '  BuyOutRangeX = "XX"
    End If
    If LevelX = 3 Then
        Me.SoldDateStart = ""
        Me.SoldDateEnd = ""
        SoldDateX = "XX"
    End If
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub WebLink_Click()
If Me.InspectionLinkWeb & "test" = "test" Then Exit Sub
Dim WebX As String
WebX = Split(Me.InspectionLinkWeb, "#")(1)
WebX = Split(WebX, "#")(0)
    Application.FollowHyperlink WebX, , True
    'GetUserAddress = True
End Sub
Private Sub LocalLink_Click()
If Me.InspectionLinkLocal & "test" = "test" Then Exit Sub
Dim WebX As String
WebX = Split(Me.InspectionLinkLocal, "#")(1)
WebX = Split(WebX, "#")(0)
    Application.FollowHyperlink WebX, , True
'    GetUserAddress = True
End Sub
