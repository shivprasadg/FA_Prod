Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim XLS, SqlA, SQLx, strSQL, TargetX, SourceCompanyX, YearX, UnitTypeX, MakeX, ModelX, BuyerX, StatusX, OrderByx As String
Dim SQLz As String
Dim SoldDateRangeX, BuyOutRangeX, PayOffSentRangeX, CityX, StateX, PayOffBackRangeX, SUFDateRangeX, AssigneeX As String
Dim intRow, i, Countx, UnitCountx As Integer
Dim RST As DAO.Recordset



Private Sub ExportToExcel_Click()
 Dim XLS, Sqlr, TitleX, UnitIDx As String
    Dim i      As Integer
    TitleX = "Admin Report " & Date
    Sqlr = "Select * From RemarketingDashboard_Admin "
    Sqlr = Sqlr & " Where AssetID in ("

    Dim RST As Recordset, intI As Integer
    Dim Rec    As Record
    Set RST = Me.RecordsetClone
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        UnitIDx = UnitIDx & RST!UnitID & ","
        RST.MoveNext
    Next i
    UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
    Sqlr = Sqlr & UnitIDx & ")"
    ' Debug.Print UnitIDx
    XLS = "|A,TL|D,TX|H,TC|G,BC|P,BC|Y,BC|AA,$E|AB,$E|AC,$E|AD,$E|AE,$E|AF,$E|AK,TC|"
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

Private Sub LocCity_Click()
    Me.LocCity.ControlTipText = Nz(DLookup("LocationAddress", "RemarketingReportAllUnits", "UnitId=" & Me.txtCurrent & ""), "No Address Listed")
End Sub

Private Sub Odometer_Click()
    Me.Odometer.ControlTipText = "ODO mileage collected from " & DLookup("ODOSource", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID) & " on " & DLookup("ODOReadDate", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID)
End Sub

Private Sub OpenMassMod_Click()
    DoCmd.close acForm, "RemarketingMassMod"
    DoCmd.OpenForm "RemarketingMassMod"
    Forms!RemarketingMassMod.ClientList = DLookup("ClientID", "vw_SixKeys", "UnitId=" & Me.txtCurrent & "")
    Form_RemarketingMassMod.ClientList_Click
    Forms!RemarketingMassMod.LeaseList = DLookup("MLID", "vw_SixKeys", "UnitId=" & Me.txtCurrent & "")
    Form_RemarketingMassMod.Leaselist_Click
    Forms!RemarketingMassMod.ScheduleList = DLookup("SchID", "vw_SixKeys", "UnitId=" & Me.txtCurrent & "")
    Form_RemarketingMassMod.ScheduleList_Click
    Forms!RemarketingMassMod.GroupList = DLookup("SGrpID", "vw_SixKeys", "UnitId=" & Me.txtCurrent & "")
    Form_RemarketingMassMod.GroupList_Click
    'Form_RemarketingMassMod_Form.Units_List_Click
    Form_RemarketingMassMod.GetAll_Click
    'Forms!RemarketingMassMod.Units_List = DLookup("UnitID", "vw_SixKeys", "UnitId=" & Me.txtCurrent & "")
End Sub

Private Sub ReportsList_Click()
    Me.ReportsList.ControlTipText = Nz(ReportsList.Column(1), "test")
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
    If Me.TargetValueList = 3 Then TargetX = " AND (Not Coalesce(PayOffRequested))"
    If Me.TargetValueList = 4 Then TargetX = " AND (Coalesce(PayOffRequested))"
    If Me.TargetValueList = 5 Then TargetX = " AND (Not Coalesce(PayOffReceived))" '
    If Me.TargetValueList = 6 Then TargetX = " AND (Coalesce(PayOffReceived))"
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
    Call ClearList("CityList")
    Call ClearList("StateList")
    Call ClearList("AssigneeList")
    ClearDates (0)
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
Private Sub ClearBank_Click()
    Call ClearList("AssigneeList")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearStatus_Click()
    Call ClearList("lstStatus")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearCity_Click()
    Call ClearList("CityList")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearState_Click()
    Call ClearList("StateList")
    Call cmdSetFilters_Click
End Sub
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
    PayOffBackRangeX = "XX"
    SUFDateRangeX = "XX"
    PayOffSentRangeX = "XX"
    OrderByx = "ORDER BY SourceCompany"
    CityX = "XX"
    StateX = "XX"
    AssigneeX = "XX"
End Sub

Public Sub SETSQL()
    SqlA = "SELECT DISTINCT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null "
    SQLz = SqlA & SourceCompanyX & YearX & UnitTypeX & MakeX & ModelX & BuyerX & StatusX & " "
    If TargetX <> "XX" Then SQLz = SQLz & TargetX
    If SoldDateRangeX <> "XX" Then SQLz = SQLz & SoldDateRangeX
    If BuyOutRangeX <> "XX" Then SQLz = SQLz & BuyOutRangeX
    If PayOffSentRangeX <> "XX" Then SQLz = SQLz & PayOffSentRangeX
    If SUFDateRangeX <> "XX" Then SQLz = SQLz & SUFDateRangeX
    If PayOffBackRangeX <> "XX" Then SQLz = SQLz & PayOffBackRangeX
    If CityX <> "XX" Then SQLz = SQLz & CityX
    If StateX <> "XX" Then SQLz = SQLz & StateX
    If AssigneeX <> "XX" Then SQLz = SQLz & AssigneeX
    If BuyerX = "" Then OrderByx = "Order By SourceCompany" Else OrderByx = "ORDER BY Buyer"
    SQLz = SQLz & OrderByx & ";"
    Dim WhereX As String
    WhereX = Split(SQLz, "RemarketingReportAllUnits ")(1)
    WhereX = Split(WhereX, "Order")(0)

    Call GetSelectedCount("CityList")
    Me.CityList.RowSource = "SELECT PCity From RemarketingReportAllUnits " & WhereX & " GROUP BY PCity ORDER BY PCity;"
    If Countx > 0 Then Call ReSelectItems("CityList")

    Call GetSelectedCount("StateList")
    Me.StateList.RowSource = "SELECT Pstate From RemarketingReportAllUnits " & WhereX & " GROUP BY Pstate ORDER BY Pstate;"
    If Countx > 0 Then Call ReSelectItems("StateList")

    Call GetSelectedCount("AssigneeList")
    Me.AssigneeList.RowSource = "SELECT BankPartner From RemarketingReportAllUnits " & WhereX & " GROUP BY BankPartner ORDER BY BankPartner;"
    If Countx > 0 Then Call ReSelectItems("AssigneeList")

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
    Me.lstModel.RowSource = "SELECT UnitModel From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitModel ORDER BY UnitModel;"
    If Countx > 0 Then Call ReSelectItems("lstModel")

    Call GetSelectedCount("lstModelYear")
    Me.lstModelYear.RowSource = "SELECT UnitYr From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitYr ORDER BY UnitYr desc;"
    If Countx > 0 Then Call ReSelectItems("lstModelYear")

    Call GetSelectedCount("lstBuyer")
    Me.lstBuyer.RowSource = "SELECT BuyerName,BuyerId From RemarketingReportAllUnits " & WhereX & " GROUP BY BuyerName, BuyerId ORDER BY BuyerName;"
    If Countx > 0 Then Call ReSelectItems("lstBuyer")
    'Me.lstStatus.RowSource = "SELECT UnitStatus From RemarketingReportAllUnits " & WhereX & " GROUP BY UnitStatus ORDER BY UnitStatus;"
    Me.Form.RecordSource = SQLz
    Me.UnitCount.Requery
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
        bName = "Buy" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 2)
        bName = "Sent" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 2)
        bName = "SUF" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 2)
        bName = "Sold" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 2)
        bName = "Back" & x
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
    Call SUFRecvStart_AfterUpdate
    Call BuyOutStart_AfterUpdate
    Call PayOffSentStart_AfterUpdate
    Call PayOffBackStart_AfterUpdate
    Call SoldDateStart_AfterUpdate
    Call CityBuild
    Call StateBuild
    Call AssigneeBuild
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
    If Coalesce(Me.lstSourceCompany, "") = "" Then
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
    SourceCompanyX = strSQL
End Sub
Private Sub ModelYearBuild()
    If Coalesce(Me.lstModelYear, "") = "" Then
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
    YearX = strSQL
End Sub
Private Sub UnitTypeBuild()
    If Coalesce(Me.lstUnitType, "") = "" Then
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
    If Coalesce(Me.lstMake, "") = "" Then
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
    If Coalesce(Me.lstModel, "") = "" Then
        strSQL = " AND ("
        With Me.lstModel
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
    If Coalesce(lstBuyer, "") = "" Then
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
    If Coalesce(Me.lstStatus, "") = "" Then
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
        strSQL = " AND ([UnitStatus] <> '') "
    End If
    StatusX = strSQL
End Sub
Private Sub CityBuild()
    If Coalesce(Me.CityList, "") = "" Then
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
Private Sub AssigneeBuild()
    If Coalesce(Me.AssigneeList, "") = "" Then
        strSQL = " AND ("
        With Me.AssigneeList
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([BankPartner]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = "XX"
        ' strSQL = " AND ([BankPartner] <> '')"
    End If
    AssigneeX = strSQL
End Sub
Private Sub StateBuild()
    If Coalesce(Me.StateList, "") = "" Then
        strSQL = " AND ("
        With Me.StateList
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([Pstate]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = "XX"
        ' strSQL = " AND ([Pstate] <> '')"
    End If
    StateX = strSQL
End Sub
Private Sub Buyer_DblClick(Cancel As Integer)
    If Nz(Me.Buyer, "") & "TEST" = "TEST" Then Exit Sub
    Call OpenBuyersFrom(Me.BuyerID, 0)
End Sub
Private Sub SUF0_Click()
    Call SUFRecvStart_Click
    Call SetDateRange("SUFRecvStart", "SUFRecvEnd", 0)
End Sub
Private Sub SUF1_Click()
    Call SUFRecvStart_Click
    Call SetDateRange("SUFRecvStart", "SUFRecvEnd", 1)
End Sub
Private Sub SUF2_Click()
    Call SUFRecvStart_Click
    Call SetDateRange("SUFRecvStart", "SUFRecvEnd", 2)
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
Private Sub Back0_Click()
    Call PayOffBackStart_Click
    Call SetDateRange("PayOffBackStart", "PayOffBackEnd", 0)
End Sub
Private Sub Back1_Click()
    Call PayOffBackStart_Click
    Call SetDateRange("PayOffBackStart", "PayOffBackEnd", 1)
End Sub
Private Sub Back2_Click()
    Call PayOffBackStart_Click
    Call SetDateRange("PayOffBackStart", "PayOffBackEnd", 2)
End Sub
Private Sub Sent0_Click()
    Call PayOffSentStart_Click
    Call SetDateRange("PayOffSentStart", "PayOffSentEnd", 0)
End Sub
Private Sub Sent1_Click()
    Call PayOffSentStart_Click
    Call SetDateRange("PayOffSentStart", "PayOffSentEnd", 1)
End Sub
Private Sub Sent2_Click()
    Call PayOffSentStart_Click
    Call SetDateRange("PayOffSentStart", "PayOffSentEnd", 2)
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
    If StartField Like "*Sent*" Then Call PayOffSentStart_AfterUpdate
    If StartField Like "*Buy*" Then Call BuyOutStart_AfterUpdate
    If StartField Like "*SUF*" Then Call SUFRecvStart_AfterUpdate
    If StartField Like "*Back*" Then Call PayOffBackStart_AfterUpdate
End Sub
Private Sub SourceCompany_DblClick(Cancel As Integer)
    If Me.SourceCompany & "TEST" = "TEST" Then Exit Sub
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
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
    Me.Buyer.SetFocus
End Sub
Private Sub ReportsList_DblClick(Cancel As Integer)
If Coalesce(Me.UnitsShown, 0) < 1 Then
MsgBox "You're Tring To Run A Report With NO Units Showing." & vbNewLine & "Please Change The Filters or Reset the Form", vbCritical, "No Shirts No Shoes, No Service"
Exit Sub
End If
    Dim FormX, CommandX As String
    FormX = Me.ReportsList.Column(3)
    CommandX = Me.ReportsList.Column(2)
    If CommandX Like "*JF*" And Me.SoldDateStart & "test" = "test" Then
    MsgBox "In order to run the JF Reports," & vbNewLine & "You must set a 'Invoice Paid' date range" & vbNewLine & "ATLAAS Core will use that date range for the various reports.", vbInformation, "Set Invoice Paid dates"
    Exit Sub
    End If
    If FormX = "Local" Then
        CallByName Me, CommandX, VbMethod
    Else
        CallByName FormX, CommandX, VbMethod
    End If
End Sub
Public Sub JFReports()
    Dim XLS, FormulaX As String
    Dim UnitCountx As Integer
    MessageUser "Open", "Creating Excel Inventory Report", "...Getting Data..."
    Call cmdSetFilters_Click
Inv_Added_As_Of_Report:
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND BuyOutDate Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# AND UnitStatus not like '*T*' Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating Inventory As Of Date Report", "...This will take just a moment..."
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,$R|AA,ZZ|AA,ZZ|AA,ZZ|AB,ZZ|AB,ZZ|"
    SendReport2Excel SqlA, "Inv_Added_Date", "Remarketing Inv Added During " & SoldDateStart & "~" & SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, False, 1
    MessageUser "Close"

Inventory_At_Date:
MessageUser "Open", "Creating Excel Inventory Report", "...Getting Data..."
    Call cmdSetFilters_Click
    'FormulaX = "=IF(OFFSET(Y#,0,1)>0,""Y"",""N"")"
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportInventory")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND (Coalesce(InvoicePaid) or (InvoicePaid > #" & Me.SoldDateEnd & "# AND InvoicePaid < #" & Me.SoldDateStart & "#)) AND UnitStatus not like '*T*' Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    Set RST = Nothing
    MessageUser "Update", "Creating Inventory @ Date Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"
    SendReport2Excel SqlA, "Inventory_@_Date", "Remarketing Inventory at " & Me.SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, False, 2, False

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
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    Set RST = Nothing
MessageUser "Update", "Creating TR Inventory Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"
    SendReport2Excel SqlA, "TR_Units", "Remarketing TR Units | " & Date & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, False, 3, False


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
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    Set RST = Nothing
MessageUser "Update", "Getting Current Inventory Report", "...This will take just a moment..."
    '|Y,Fx" & FormulaX & "
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|L,BC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|T,$E|U,$E|V,$E|W,$E|Y,BC|Z,P2|"
    SendReport2Excel SqlA, "Current_Inventory", "Remarketing Inventory at " & Date & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, False, 3, False

ExportSoldReport:
    SqlA = SQLz
    SqlA = Replace(SqlA, "RemarketingReportAllUnits", "RemarketingReportSoldUnits")
    SqlA = Replace(SqlA, "SourceCompany", "[Account Name]")
    SqlA = Split(SqlA, "WHERE")(0)
    SqlA = SqlA & " WHERE [UNITId] is not null AND InvoicePaid Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# Order By [Account Name];"
    Set RST = CurrentDb.OpenRecordset(SqlA)
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    Set RST = Nothing
MessageUser "Update", "Creating Excel Sold Report", "...This will take just a moment..."
    XLS = "|E,TX|E,BC|F,TC|I,BL|J,SC|M,TC|N,$E|O,$E|P,$E|Q,$E|R,$E|S,$E|W,BC|Z,ZZ|AA,ZZ|AA,ZZ|"
    SendReport2Excel SqlA, "Sold_Units", "Remarketing Sold Units " & SoldDateStart & "~" & SoldDateEnd & " | " & Format(UnitCountx, "#,#") & " Units ", True, XLS, UnitCountx, False, 4, True
MessageUser "Close"

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
Public Sub RVRealization()
    Dim DateFilter As String
    Dim DateRange As String
    Dim TitleX As String
    Dim AndWhereX As String ' Subset of the Forms' SQL Statement

    DateFilter = "DatePart('yyyy',BuyOutDate) AS BuyOut"

    If IsDate(Me.BuyOutStart) Then
        DateRange = Me.BuyOutStart & "~" & Me.BuyOutEnd
    End If
    If IsDate(Me.SoldDateStart) Then
        DateFilter = "DatePart('yyyy',InvoicePaid) AS Sold"
        DateRange = Me.SoldDateStart & "~" & Me.SoldDateEnd
    End If
   ' If IsDate(Me.InvoiceSentStart) Then
   '     DateFilter = "DatePart('yyyy',InvoiceSent) AS Inv"
   '     DateRange = Me.InvoiceSentStart & "~" & Me.InvoiceSentEnd
   ' End If
    If IsDate(Me.SUFRecvStart) Then
        DateFilter = "DatePart('yyyy',SUFRecv) AS SUF"
        DateRange = Me.SUFRecvStart & "~" & Me.SUFRecvEnd
    End If
    If Me.Form.RecordSource Like "*AND*" Then
        AndWhereX = SplitString(SplitString(Me.Form.RecordSource, "is not null ", 1), "Order By", 0)
    Else
        AndWhereX = ""
    End If
    TitleX = Date & " - By " & SplitString(DateFilter, "AS ", 1)
    If DateRange <> "" Then TitleX = TitleX & "( " & DateRange & " )"
    If AndWhereX Like "*Bank*" Then TitleX = TitleX & ", Assignee"
    If AndWhereX Like "*Client*" Then TitleX = TitleX & ", Client"

    Call RVRealizationReport(TitleX, DateFilter, AndWhereX)

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
Private Sub PayOffSentStart_AfterUpdate()
    If IsDate(Me.PayOffSentStart) And IsDate(Me.PayOffSentEnd) Then
        PayOffSentRangeX = "And PayOffRequested Between #" & Me.PayOffSentStart & "# AND #" & Me.PayOffSentEnd & "# "
        OrderByx = "ORDER BY PayOffRequested Desc"
        '    Call SETSQL
    Else
        PayOffSentRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
End Sub
Private Sub PayOffBackStart_AfterUpdate()
    If IsDate(Me.PayOffBackStart) And IsDate(Me.PayOffBackEnd) Then
        PayOffBackRangeX = "And PayOffReceived Between #" & Me.PayOffBackStart & "# AND #" & Me.PayOffBackEnd & "# "
        OrderByx = "ORDER BY PayOffReceived Desc"
        '    Call SETSQL
    Else
        PayOffBackRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
End Sub
Private Sub PayOffSentEnd_AfterUpdate()
    Call PayOffSentStart_AfterUpdate
End Sub
Private Sub PayOffBackEnd_AfterUpdate()
    Call PayOffBackStart_AfterUpdate
End Sub
Private Sub BuyOutStart_AfterUpdate()
    If IsDate(Me.BuyOutEnd) And IsDate(Me.BuyOutStart) Then
        BuyOutRangeX = "And BuyOutDate Between #" & Me.BuyOutStart & "# AND #" & Me.BuyOutEnd & "# "
        OrderByx = "ORDER BY BuyOutDate Desc"
        'Call SETSQL
    Else
        BuyOutRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
    'If Not IsDate(Me.BuyOutEnd) And Not IsDate(Me.BuyOutStart) Then Call SETSQL
End Sub
Private Sub BuyOutEnd_AfterUpdate()
    Call BuyOutStart_AfterUpdate
End Sub
Private Sub SoldDateStart_AfterUpdate()
    If IsDate(Me.SoldDateEnd) And IsDate(Me.SoldDateStart) Then
        SoldDateRangeX = "And InvoicePaid Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# "
        OrderByx = "ORDER BY InvoicePaid Desc"
        '  Call SETSQL
    Else
        SoldDateRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
    ' If Not IsDate(Me.SoldDateEnd) And Not IsDate(Me.SoldDateStart) Then Call SETSQL
End Sub
Private Sub SoldDateEnd_AfterUpdate()
    Call SoldDateStart_AfterUpdate
End Sub
Private Sub SUFRecvStart_AfterUpdate()
    If IsDate(Me.SUFRecvEnd) And IsDate(Me.SUFRecvStart) Then
        SUFDateRangeX = "And [SUF Returned] Between #" & Me.SUFRecvStart & "# AND #" & Me.SUFRecvEnd & "# "
        OrderByx = "ORDER BY [SUF Returned] Desc"
        'Call SETSQL
    Else
        SUFDateRangeX = "XX"
        OrderByx = "ORDER BY SourceCompany"
    End If
    'If Not IsDate(Me.BuyOutEnd) And Not IsDate(Me.BuyOutStart) Then Call SETSQL
End Sub
Private Sub SUFRecvEnd_AfterUpdate()
    Call SUFRecvStart_AfterUpdate
End Sub

Private Sub PayOffSentStart_Click()
    ClearDates (1)
End Sub
Private Sub PayOffSentEnd_Click()
    ClearDates (1)
End Sub
Private Sub BuyOutStart_Click()
    ClearDates (2)
End Sub
Private Sub BuyOutEnd_Click()
    ClearDates (2)
End Sub
Private Sub SoldDateStart_Click()
    ClearDates (3)
End Sub
Private Sub SoldDateEnd_Click()
    ClearDates (3)
End Sub
Private Sub SUFRecvStart_Click()
    ClearDates (4)
End Sub
Private Sub SUFRecvEnd_Click()
    ClearDates (4)
End Sub
Private Sub PayOffBackStart_Click()
    ClearDates (5)
End Sub
Private Sub PayOffBackEnd_Click()
    ClearDates (5)
End Sub


Private Sub ClearDates(LevelX As Integer)
    If LevelX <> 1 Then
        Me.PayOffSentStart = ""
        Me.PayOffSentEnd = ""
        PayOffSentRangeX = "XX"
    End If
    If LevelX <> 2 Then
        Me.BuyOutStart = ""
        Me.BuyOutEnd = ""
        BuyOutRangeX = "XX"
    End If
    If LevelX <> 3 Then
        Me.SoldDateStart = ""
        Me.SoldDateEnd = ""
        SoldDateX = "XX"
    End If
    If LevelX <> 4 Then
        Me.SUFRecvStart = ""
        Me.SUFRecvEnd = ""
        SUFDateRangeX = "XX"
    End If
    If LevelX <> 5 Then
        Me.PayOffBackStart = ""
        Me.PayOffBackEnd = ""
        PayOffBackRangeX = "XX"
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
