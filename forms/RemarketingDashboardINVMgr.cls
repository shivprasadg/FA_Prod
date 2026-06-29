Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim XLS, SqlA, SQLx, strSQL, TargetX, SourceCompanyX, YearX, UnitTypeX, MakeX, ModelX, BuyerX As String
Dim SQLz As String
Dim StatusX, OrderByx, SoldDateRangeX, BuyOutRangeX, InvoiceSentRangeX, CityX, StateX, AssigneeX As String
Dim intRow, i, Countx, UnitCountx As Integer
Dim RST As DAO.Recordset

Private Sub ExportToExcel_Click()
    Dim XLS, TitleX, Sqlr, UnitIDx As String
    Dim i      As Integer
    TitleX = "Inventory Report Mgr"
    Sqlr = "Select SourceCompany As [Source Company]" 'A
    Sqlr = Sqlr & ",UnitID As AssetID"                'B
    Sqlr = Sqlr & ",Unitnum "                         'C
    Sqlr = Sqlr & ",TargetSalePrice AS [Target SalePrice]"  'D
    Sqlr = Sqlr & ",UnitStatus AS [Unit Status] "     'E
    Sqlr = Sqlr & ",RStatus as [Inv Status]"          'F
    Sqlr = Sqlr & ",UnitVIN As VIN"                   'G
    Sqlr = Sqlr & ",UnitYr As Myear"                  'H
    Sqlr = Sqlr & ",UnitMake As Make"                 'I
    Sqlr = Sqlr & ",UnitModel As Model"               'J
    Sqlr = Sqlr & ",UnitSubType As SubType"           'K
    Sqlr = Sqlr & ",Axle "                            'L
    Sqlr = Sqlr & ",Engine "                          'M
    Sqlr = Sqlr & ",Trans "                           'N
    Sqlr = Sqlr & ",Locationsn as Domicile"           'O
    Sqlr = Sqlr & ",ParkedLocation As ParkedLoc "     'P
    Sqlr = Sqlr & ",[Move From] "                     'Q
    Sqlr = Sqlr & ",[Move Start] "                    'R
    Sqlr = Sqlr & ",[Move To] "                       'S
    Sqlr = Sqlr & ",PState "                          'T
    Sqlr = Sqlr & ",[Sch|Grp] As [Lease Sch|Grp]"     'U
    Sqlr = Sqlr & ",InspectionLinkWeb As InspWeb"     'V
    Sqlr = Sqlr & ",InspectionPassed As [Insp Passed]"   'W
    Sqlr = Sqlr & ",OffLeaseDate As [Off Lease Actual]"  'X
    Sqlr = Sqlr & ",DaysInInventory As [Days IN]"     'Y
    Sqlr = Sqlr & ",CurrBook "                        'Z
    Sqlr = Sqlr & ",BuyerType As Disp"                'AA
    Sqlr = Sqlr & ",BuyerName As Buyer"               'AB
    Sqlr = Sqlr & ",UnitSalePrice As [Sold Price]"    'AC
    Sqlr = Sqlr & ",InvoicePaid As [Sold Date]"       'AD
    Sqlr = Sqlr & ",BankPartner as Bank "             'AE
    Sqlr = Sqlr & ",IIF(BankShareYN='Y','Share','Exchange') as [Bank Share]" 'AF
    Sqlr = Sqlr & ",TotalExpense as RMKT_Exp "        'AG
    Sqlr = Sqlr & ",BuyOut "                          'AH
    Sqlr = Sqlr & ",Residual As RV"                   'AI
    Sqlr = Sqlr & ",ODORead As ODO"                   'AJ
    Sqlr = Sqlr & ",ODOSource AS [ODO Source]"        'AK
    Sqlr = Sqlr & ",ODOReadDate As [ODO Date]"        'AL
    Sqlr = Sqlr & ",LeaseExp As [Lease Expire]"       'AM
    Sqlr = Sqlr & ",ProjEOL as [Projected EOL]"       'AN
    Sqlr = Sqlr & " From RemarketingReportAllUnits "
    Sqlr = Sqlr & " Where UnitID in ("

    Dim RST As Recordset
    Set RST = Me.RecordsetClone
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        UnitIDx = UnitIDx & RST!UnitID & ","
        RST.MoveNext
    Next i
    UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
    Sqlr = Sqlr & UnitIDx & ") ORDER BY SourceCompany, UnitID "

    XLS = "|A,DD3|A,Wrap|A,BC|B,TX|B,TC|C,TC|D,$R|E,BC|H,TC|"
    XLS = XLS & "L,TC|V,TC|R,TC|V,HY|T,TC|X,TC|Y,TC|"
    XLS = XLS & "Z,$R|AA,TC|AC,$R|AD,$R|AF,TC|AG,$E|AH,$E|AI,$R|AJ,SIntR|AK,TC|AM,TC|AN,TC|"
    XLS = XLS & "B,W10|D,W12|E,W06|P,W20|Q,W20|R,W11|S,W20|V,W07|X,W10|Y,W06|AK,W07|AF,W08|AM,W10|AN,W10|"
    XLS = XLS & "AE,cColor:3:green:End=AF|AJ,cColor:3:orange:End=AL|"
    XLS = XLS & "AE,cBorder:Med:End=AF|AJ,cBorder:Med:End=AL|"

    MessageUser "Open", "Getting Data...", "This will be quick"
    SendReport2Excel Sqlr, "Inventory", TitleX, True, XLS, 0, False, , , , "RMKTInventoryReportMgR"
    MessageUser "Close"
End Sub

Private Sub Form_Load()
    Me.Form.FilterOn = False
    Call SetYearList
    Call ClearBoxes
  '  Call SetButtonYears
    Call cmdSetFilters_Click
End Sub

Private Sub LocCity_Click()
Me.LocCity.ControlTipText = DLookup("LAddrLine1", "Locations", "LocationID=" & Me.LocationID & "")
End Sub

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
Public Sub RVRealization()
    Dim DateFilter As String
    Dim DateRange As String
    Dim TitleX As String
    Dim AndWhereX As String ' Subset of the Forms' SQL Statement

    DateFilter = "DatePart('yyyy',InvoicePaid) AS Sold"
   DateRange = (Date - 365) & "~" & Date

   DateRange = Date
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
    Call ClearList("CityList")
    Call ClearList("StateList")
'    ClearDates (1)
'    ClearDates (2)
'    ClearDates (3)
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
    StatusX = "AND UnitStatus is not null "
    SoldDateRangeX = "XX"
    BuyOutRangeX = "XX"
    InvoiceSentRangeX = "XX"
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
    'If InvoiceSentRangeX <> "XX" Then SQLz = SQLz & InvoiceSentRangeX
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
End Sub

Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Group1" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
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
        strSQL = " AND ([UnitStatus] <> 'A' AND [UnitStatus] <> 'B' AND [UnitStatus] <> 'T') "
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
    'DoCmd.close acForm, "UnitsForm"
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UnitId=" & Me.UnitID & ""
    Me.UnitDesc.SetFocus
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
Public Sub JFReports()
MsgBox "This report is not avail from this form", , AtlaasVersion

End Sub

'Private Sub InvoiceSentEnd_AfterUpdate()
'    Call InvoiceSentStart_AfterUpdate
'End Sub
'
'Private Sub BuyOutEnd_AfterUpdate()
'    Call BuyOutStart_AfterUpdate
'End Sub
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
'
'Private Sub SoldDateStart_Click()
'    ClearDates (1)
'    ClearDates (2)
'End Sub
'Private Sub SoldDateEnd_Click()
'    ClearDates (1)
'    ClearDates (2)
'End Sub
'Private Sub InvoiceSentStart_Click()
'    ClearDates (3)
'    ClearDates (2)
'End Sub
'Private Sub InvoiceSentEnd_Click()
'    ClearDates (3)
'    ClearDates (2)
'End Sub
'Private Sub BuyOutStart_Click()
'    ClearDates (1)
'    ClearDates (3)
'End Sub
'Private Sub BuyOutEnd_Click()
'    ClearDates (1)
'    ClearDates (3)
'End Sub
'
'Private Sub ClearDates(LevelX As Integer)
'    If LevelX = 1 Then
'        '   Me.InvoiceSentStart = ""
'        '   Me.InvoiceSentEnd = ""
'        '   InvoiceSentRangeX = "XX"
'    End If
'    If LevelX = 2 Then
'        '  Me.BuyOutStart = ""
'        '  Me.BuyOutEnd = ""
'        '  BuyOutRangeX = "XX"
'    End If
'    If LevelX = 3 Then
'        'Me.SoldDateStart = ""
'        'Me.SoldDateEnd = ""
'        'SoldDateX = "XX"
'    End If
'End Sub

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
