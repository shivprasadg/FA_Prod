Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim XLS, SqlA, strSQL, TargetX, SourceCompanyX, YearX, UnitTypeX, MakeX, ModelX, BuyerX, StatusX, OrderByx, SoldDateRangeX, BuyOutRangeX, InvoiceSentRangeX, CityX, StateX As String
Dim SQLz As String
Dim SQLx As String
Dim SQLxz As String
Dim intRow, i, Countx, UnitCountx As Integer
Dim RST        As DAO.Recordset
Private Sub Form_Load()
    Me.Form.FilterOn = False
    Call SetYearList
    Call ClearBoxes
    Call SetButtonYears
    Call cmdSetFilters_Click
End Sub
Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub
Private Sub ExportToExcel_Click()
    Dim XLS, QTRx, TitleX, Sqlr, UnitIDx As String
    Dim i      As Integer
    TitleX = "Inventory Report"                                 'Col A
    Sqlr = "Select SourceCompany "                              'Col B
    Sqlr = Sqlr & ",UnitID As AssetID"                          'Col C
    Sqlr = Sqlr & ",Unitnum "                                   'Col D
    Sqlr = Sqlr & ",TargetSalePrice "                           'Col E
    Sqlr = Sqlr & ",UnitStatus "                                'Col F
    Sqlr = Sqlr & ",RStatus aS InvStatus"                       'Col G
    Sqlr = Sqlr & ",UnitVIN As VIN"                             'Col H
    Sqlr = Sqlr & ",UnitYr As Myear"                            'Col I
    Sqlr = Sqlr & ",UnitMake As Make"                           'Col J
    Sqlr = Sqlr & ",UnitModel As Model"                         'Col K
    Sqlr = Sqlr & ",ODORead As ODO"                             'Col L
    Sqlr = Sqlr & ",UnitSubType As SubType"                     'Col M
    Sqlr = Sqlr & ",Axle "                                      'Col N
    Sqlr = Sqlr & ",Engine "                                    'Col O
    Sqlr = Sqlr & ",Trans "                                     'Col P
    Sqlr = Sqlr & ",'' as [FCAM Y/N] "                          'Col Q
    Sqlr = Sqlr & ",'' as [SLPR Bunks] "                        'Col R
    Sqlr = Sqlr & ",ParkedLocation "                            'Col S
    Sqlr = Sqlr & ",ParkedAddress "                             'Col T
    Sqlr = Sqlr & ",PCity "                                     'Col U
    Sqlr = Sqlr & ",PState "                                    'Col V
    Sqlr = Sqlr & ",PZip "                                      'Col W
    Sqlr = Sqlr & ",InspectionLinkWeb As InspWeb"               'Col X
    Sqlr = Sqlr & ",InspectionPassed As InspPassed"             'Col Y
    Sqlr = Sqlr & ",OffLeaseDate As OffLease"                   'Col Z
    Sqlr = Sqlr & ",DaysInInventory As [Days IN]"               'Col AA
    Sqlr = Sqlr & ",MLOrig "                                    'Col AB
    Sqlr = Sqlr & ",[Sch|Grp] As [Lease Sch|Grp]"               'Col AC
    Sqlr = Sqlr & ",BuyerType As Disp"                          'Col AD
    Sqlr = Sqlr & ",BuyerName As Buyer"                         'Col AE
    Sqlr = Sqlr & ",UnitSalePrice As SoldPrice"                 'Col AF
    Sqlr = Sqlr & ",InvoicePaid As SoldDate"                    'Col AG
    Sqlr = Sqlr & ",BankShareYN "                               'Col AH
    Sqlr = Sqlr & ",TotalExpense as RMKT_Exp "                  'Col AI
    Sqlr = Sqlr & ",ODOSource "                                 'Col AJ
    Sqlr = Sqlr & ",ODOReadDate As ODODate"                     'Col AK
    Sqlr = Sqlr & ",Residual As RV"                             'Col AL
    Sqlr = Sqlr & ",ProjEOL "                                   'Col AM
    Sqlr = Sqlr & ",ParkDateProjected As ParkedProj"            'Col AN

    Sqlr = Sqlr & " From RemarketingReportAllUnits "
    Sqlr = Sqlr & " Where UnitID in ("

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
    XLS = "|A,BC|B,TX|B,TR|C,TC|D,$R|E,BC|H,TC|K,SIntR|W,TC|W,HY|AA,TC|AD,TC|AF,$R|AH,TC|AI,$E|AL,$R|"
    MessageUser "Open", "Getting Data...", "This will be quick"
    SendReport2Excel Sqlr, "Inventory", "Fleet Advantage: Virtual Used Truck Factory", True, XLS, 0, False, , , , "RMKTInventoryReport"
    MessageUser "Close"
End Sub

'Private Sub FindAllABuyer_Click()
'    If IsNull(SQLz) Then
'        Call cmdSetFilters_Click
'    End If
'    Set RST = CurrentDb.OpenRecordset(SQLz)
'    RST.MoveFirst
'    RST.MoveLast
'    UnitCountx = RST.RecordCount
'    RST.MoveFirst
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL ("DELETE * From RemarketingSales_WorkTable;")
'    For i = 1 To UnitCountx
'        DoCmd.RunSQL ("INSERT INTO RemarketingSales_WorkTable ( UnitID ) SELECT " & RST!UnitID & ";")
'        RST.MoveNext
'    Next i
'    DoCmd.SetWarnings True
'    Set RST = Nothing
'    DoCmd.close acForm, "RemarketingSalesPersonsDashBoard"
'    Dim OpenArgX As String
'    If FAUserAccess > 9 Then OpenArgX = "Manager"
'    If FAUserAccess < 9 Then OpenArgX = "Sales"
'    If IsOpen("rmkt menu") And Forms![rmkt menu].OverRideCbo = "Sales" Then OpenArgX = "Sales"
'   ' DoCmd.OpenForm "RemarketingSalesPersonsDashBoard", , , , , , OpenArgX
'    DoCmd.close acForm, "RemarketingProposalManager"
'    DoCmd.OpenForm "RemarketingProposalManager"
'End Sub

Private Sub Odometer_Click()
    Me.Odometer.ControlTipText = "ODO mileage collected from " & DLookup("ODOSource", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID) & " on " & DLookup("ODOReadDate", "RemarketingReportAllUnits", "UnitId=" & Me.UnitID)
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
Private Sub SetYearList()
    Dim YearListX As String
    yearlist = Chr(34) & "ALL" & Chr(34) & ";"">0" & Chr(34)
    Dim x      As Integer
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
    StatusX = "AND UnitStatus Not Like '*S*' AND UnitStatus <> 'D' "
    SoldDateRangeX = "XX"
    BuyOutRangeX = "XX"
    InvoiceSentRangeX = "XX"
    OrderByx = "ORDER BY SourceCompany"
    CityX = "XX"
    StateX = "XX"
End Sub

Public Sub SETSQL()

   On Error GoTo SETSQL_Error

    SqlA = "SELECT DISTINCT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null "
    SQLz = SqlA & SourceCompanyX & YearX & UnitTypeX & MakeX & ModelX & BuyerX & StatusX & " "
    If TargetX <> "XX" Then SQLz = SQLz & TargetX
    If SoldDateRangeX <> "XX" Then SQLz = SQLz & SoldDateRangeX
    If BuyOutRangeX <> "XX" Then SQLz = SQLz & BuyOutRangeX
    'If InvoiceSentRangeX <> "XX" Then SQLz = SQLz & InvoiceSentRangeX
    If CityX <> "XX" Then SQLz = SQLz & CityX
    If StateX <> "XX" Then SQLz = SQLz & StateX
    If BuyerX = "" Then OrderByx = "Order By SourceCompany" Else OrderByx = "ORDER BY Buyer"
    SQLz = SQLz & OrderByx & ";"
    Dim WhereX As String
    WhereX = Split(SQLz, "RemarketingReportAllUnits ")(1)
    WhereX = Split(WhereX, "Order")(0)

    Call GetSelectedCount("CityList")
    Me.CityList.RowSource = "SELECT PCity From RemarketingReportAllUnits " & WhereX & " GROUP BY PCity ORDER BY PCity;"
    If Countx > 0 Then Call ReSelectItems("CityList")

    Call GetSelectedCount("StateList")
    Me.StateList.RowSource = "SELECT PState From RemarketingReportAllUnits " & WhereX & " GROUP BY PState ORDER BY PState;"
    If Countx > 0 Then Call ReSelectItems("StateList")

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

   On Error GoTo 0
   Exit Sub

SETSQL_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SETSQL of VBA Document Form_RemarketingDashboardInvSales"
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Group1" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub
Public Sub SetButtonYears()
'Dim x As Integer
'Dim bName As String
'For x = 0 To 2
'  bName = "Buy" & X
'  Me(bName).Caption = DatePart("yyyy", Date) - X
'  bName = "Sent" & X
'  Me(bName).Caption = DatePart("yyyy", Date) - X
'   bName = "Sold" & x
'   Me(bName).Caption = DatePart("yyyy", Date) - x
'Next x
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
    ' Call SoldDateStart_AfterUpdate
    ' Call InvoiceSentStart_AfterUpdate
    Call CityBuild
    Call StateBuild
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
    If Coalesce(Me.lstBuyer, "") = "" Then
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
        strSQL = " AND ([UnitStatus] <> 'S' AND [UnitStatus] <> 'D' AND [UnitStatus] <> 'KS') "
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
Private Sub StateBuild()
    If Coalesce(Me.StateList, "") = "" Then
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
    If Coalesce(Me.Buyer, "") = "" Then Exit Sub
    Call OpenBuyersFrom(Me.BuyerID, 0)
End Sub
'Private Sub Buy0_Click()
'    Call BuyOutStart_Click
'    Call SetDateRange("BuyOutStart", "BuyOutEnd", 0)
'End Sub
'Private Sub Buy1_Click()
'    Call BuyOutStart_Click
'    Call SetDateRange("BuyOutStart", "BuyOutEnd", 1)
'End Sub
'Private Sub Buy2_Click()
'    Call BuyOutStart_Click
'    Call SetDateRange("BuyOutStart", "BuyOutEnd", 2)
'End Sub
'Private Sub Sent0_Click()
'    Call InvoiceSentStart_Click
'    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 0)
'End Sub
'Private Sub Sent1_Click()
'    Call InvoiceSentStart_Click
'    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 1)
'End Sub
'Private Sub Sent2_Click()
'    Call InvoiceSentStart_Click
'    Call SetDateRange("InvoiceSentStart", "InvoiceSentEnd", 2)
'End Sub
'Private Sub Sold0_Click()
'    Call SoldDateStart_Click
'    Call SetDateRange("SoldDateStart", "SoldDateEnd", 0)
'End Sub
'Private Sub Sold1_Click()
'    Call SoldDateStart_Click
'    Call SetDateRange("SoldDateStart", "SoldDateEnd", 1)
'End Sub
'Private Sub Sold2_Click()
'    Call SoldDateStart_Click
'    Call SetDateRange("SoldDateStart", "SoldDateEnd", 2)
'End Sub
'Private Sub SetDateRange(StartField As String, EndField As String, YearBack As Integer)
'    Me(StartField) = "1/1/" & DatePart("yyyy", Date) - YearBack
'    If YearBack = 0 Then Me(EndField) = Date Else Me(EndField) = "12/31/" & DatePart("yyyy", Date) - YearBack
'    If StartField Like "*Sold*" Then Call SoldDateStart_AfterUpdate
'    If StartField Like "*Sent*" Then Call InvoiceSentStart_AfterUpdate
'    If StartField Like "*Buy*" Then Call BuyOutStart_AfterUpdate
'End Sub

Private Sub SetCity_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetCompany_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetMake_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetModel_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetState_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetStatus_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetTarget_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetType_Click()
Call cmdSetFilters_Click
End Sub
Private Sub SetYear_Click()
Call cmdSetFilters_Click
End Sub

Private Sub SourceCompany_DblClick(Cancel As Integer)
    If Me.SourceCompany & "TEST" = "TEST" Then Exit Sub
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
    Me.UnitID.SetFocus
End Sub

Private Sub SPDashboard_Click()
 'DoCmd.close acForm, "RemarketingSalesPersonsDashBoard"
    Dim OpenArgX As String
    If FAUserAccess > 9 Then OpenArgX = "Manager"
    If FAUserAccess < 9 Then OpenArgX = "Sales"
    If IsOpen("rmkt menu") And Forms![rmkt menu].OverRideCbo = "Sales" Then OpenArgX = "Sales"
    DoCmd.OpenForm "RemarketingSalesPersonsDashBoard", , , , , , OpenArgX
End Sub

Private Sub TargetValueList_Click()
    Call TargetPriceSetBuild
    'Call SETSQL
End Sub
Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UnitId=" & Me.UnitID & ""
    Me.ActiveControl.SetFocus
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

    SQLx = "Select SourceCompany as [Account Name]"             'Col A
    SQLx = SQLx & ",UnitID as AssetID"                          'Col B
    SQLx = SQLx & ",UnitNum As [Unit Num]"                      'Col C
    SQLx = SQLx & ",InspectionLinkWeb as [Insp Link (Click)]"   'Col D
    SQLx = SQLx & ",UnitStatus As [Unit Status] "               'Col E
    SQLx = SQLx & ",RStatus As [Inventory Status] "             'Col F
    SQLx = SQLx & ",UnitSalePrice as SalePrice   "              'Col G
    SQLx = SQLx & ",UnitVIN as VIN"                             'Col H
    SQLx = SQLx & ",UnitYr as Myear"                            'Col I
    SQLx = SQLx & ",UnitMake AS Make"                           'Col J
    SQLx = SQLx & ",UnitModel AS Model"                         'Col K
    SQLx = SQLx & ",UnitType as [Sub Type]"                     'Col L
    SQLx = SQLx & ",Axle "                                      'Col M
    SQLx = SQLx & ",ODORead as Miles"                           'Col N
    SQLx = SQLx & ",Engine"                                     'Col O
    SQLx = SQLx & ",'' as HP"                                   'Col P
    SQLx = SQLx & ",Trans"                                      'Col Q
    SQLx = SQLx & ",'' as FCAM"                                 'Col R
    SQLx = SQLx & ",'' as Bunk"                                 'Col S
    SQLx = SQLx & ",ParkedLocation AS [Parked Loc]"             'Col T
    SQLx = SQLx & ",PCity"                                      'Col U
    SQLx = SQLx & ",PState"                                     'Col V
    SQLx = SQLx & ",Pzip"                                       'Col W
    SQLx = SQLx & ",[Sch|Grp]"                                  'Col X
    SQLx = SQLx & ",InspectionPassed as [INSP Pass]"            'Col Y
    SQLx = SQLx & ",OffLeaseDate as [OffLease Date]"            'Col Z
    SQLx = SQLx & ",BankPartner As [Bank Partner]"              'Col AA
    SQLx = SQLx & " From RemarketingReportAllUnits"
    SQLx = SQLx & " WHERE UnitStatus IN ('EF','EG','EH','EI','EM','EP','EW ','EZ','P','PS','R') "
    SQLx = SQLx & " Order By SourceCompany"
    Dim Y As Integer
        For Y = 0 To Me.lstStatus.ListCount
            If (Me.lstStatus.ItemData(Y) = ("EF")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EG")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EH")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EI")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EM")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EP")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EW")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("EZ")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("P")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("PS")) Then Me.lstStatus.Selected(Y) = True
            If (Me.lstStatus.ItemData(Y) = ("R")) Then Me.lstStatus.Selected(Y) = True
        Next Y
    Call cmdSetFilters_Click
    'SQLx = SQLx & " WHERE " & SplitString(SQLz, "WHERE", 1) & ""
    XLS = "|A,mTable:Name=factory|A,Wrap|A,DD|A,BL|B,TX|B,BC|C,TC|D,HY|E,TC|E,BC|F,TC|G,BR|G,$R|I,TC|M,TC|V,TC|N,BR|N,Sint|D,TC|L,BC|"
    XLS = XLS & "D,cBorder:Med:End=D|T,cBorder:Med:End=W|Y,cBorder:Med:End=AA|Z,cBorder:Med|"
    XLS = XLS & "B,cColor:3:black|D,cColor:3:maroon|N,cColor:3:black|G,cColor:3:green|T,cColor:3:green:End=W|Z,cColor:3:maroon|"
    XLS = XLS & "B,W10|D,W09|E,W07|F,W11|P,W09|Z,TC|Z,W10|Y,TC|AA,TC|AA,W10|"
    SendReport2Excel SQLx, "Inventory", "Fleet Advantage: Virtual Used Truck Factory: " & Date & "", True, XLS, 0, , , , , "RMKTInventory"
'    Call Remarketing_INVReport_Sales(SQLx, 0)
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
'Private Sub SoldDateStart_AfterUpdate()
'    If IsDate(Me.SoldDateEnd) And IsDate(Me.SoldDateStart) Then
'        SoldDateRangeX = "And InvoicePaid Between #" & Me.SoldDateStart & "# AND #" & Me.SoldDateEnd & "# "
'        OrderByX = "ORDER BY InvoicePaid"
'        '  Call SETSQL
'    Else
'        SoldDateRangeX = "XX"
'        OrderByX = "ORDER BY SourceCompany"
'    End If
'    ' If Not IsDate(Me.SoldDateEnd) And Not IsDate(Me.SoldDateStart) Then Call SETSQL
'End Sub

'Private Sub SoldDateEnd_AfterUpdate()
'    SoldDateStart_AfterUpdate
'End Sub

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
        'Me.SoldDateStart = ""
        'Me.SoldDateEnd = ""
        'SoldDateX = "XX"
    End If
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub WebLink_Click()
    If Me.InspectionLinkWeb & "test" = "test" Then Exit Sub
    Dim WebX   As String
    WebX = Split(Me.InspectionLinkWeb, "#")(1)
    WebX = Split(WebX, "#")(0)
    Application.FollowHyperlink WebX, , True
    'GetUserAddress = True
End Sub
Private Sub LocalLink_Click()
    If Me.InspectionLinkLocal & "test" = "test" Then Exit Sub
    Dim WebX   As String
    WebX = Split(Me.InspectionLinkLocal, "#")(1)
    WebX = Split(WebX, "#")(0)
    Application.FollowHyperlink WebX, , True
    '    GetUserAddress = True
End Sub
