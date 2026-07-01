Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim XLS, SqlA, SQLz, SQLx, strSQL, TargetX, AccountNameX, YearX, DescX, Sch_GrpX, BDEXz As String
Dim StatusX, OrderByx, FundDateRange, BuyOutRangeX, InvoiceSentRangeX, CityX, StateX, BankX, CGroupX As String
Dim intRow, i, Countx, UnitCountx As Integer
Dim RST As DAO.Recordset

Private Sub Form_Load()
Call MakeTFRWorkTable
Me.Form.FilterOn = False
Call ClearBoxes
Call cmdSetFilters_Click
Call SetButtonYears
End Sub

Private Sub FundingStatus_DblClick(Cancel As Integer)
Dim Msgx As String
If Me.FundingStatus = "F" Then Msgx = "... then again, you already knew that :)"
If GroupFinalizedCheck(Me.GroupID) = "Good" Then MsgBox "This Group checks out 'Good'" & vbNewLine & " (you may finalize)" & Msgx, vbInformation, AtlaasVersion
End Sub

Private Sub GetData_Click()
Call MakeTFRWorkTable(True)  'moved to a function
End Sub
Public Sub SetListBoxDefaults()
    Me.lstCGroup.RowSource = "SELECT TFR.ClientGroup, TFR.ClientGroupId FROM TFR_Work AS TFR WHERE TFR.AccountName Is Not Null GROUP BY TFR.ClientGroup, TFR.ClientGroupId ORDER BY TFR.ClientGroup;"
    Me.lstAccountName.RowSource = "SELECT TFR.AccountName, ClientId FROM TFR_Work As TFR WHERE TFR.AccountName Is Not Null GROUP BY TFR.AccountName, TFR.ClientId ORDER BY TFR.AccountName;"
    Me.lstFundYear.RowSource = "SELECT TFR.FYear FROM TFR_Work As TFR WHERE TFR.[Sch_Grp] Is Not Null GROUP BY TFR.FYear ORDER BY TFR.FYear DESC;"
    Me.lstBankType.RowSource = "SELECT TFR.Bank FROM TFR_Work As TFR WHERE TFR.SGrpID Is Not Null GROUP BY TFR.Bank ORDER BY TFR.Bank;"
    Me.lstDesc.RowSource = "SELECT UnitDesc From TFR_Work WHERE ClientID is not null   GROUP BY UnitDesc ORDER BY UnitDesc;"
    Me.lstSch_Grp.RowSource = "SELECT Sch_Grp From TFR_Work WHERE ClientID is not null   GROUP BY Sch_Grp ORDER BY Sch_Grp;"
    Me.lstStatus.RowSource = "SELECT TFR.Assigned_Indicator FROM TFR_Work as TFR WHERE TFR.Assigned_Indicator <>'X' GROUP BY TFR.Assigned_Indicator ORDER BY TFR.Assigned_Indicator;"
    Me.Form.RecordSource = "Select * From TFR_Work"
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
    Dim x As Integer
    For x = CInt(DatePart("yyyy", Date) + 1) To CInt(DatePart("yyyy", Date) - 12) Step -1
        yearlist = yearlist & ";" & Chr(34) & x & Chr(34) & ";" & Chr(34) & "=" & x & Chr(34)
    Next x
    Me.lstFundYear.RowSource = yearlist
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.GroupID
    On Error GoTo 0
End Sub

Private Sub ClearAll_Click()
    Call ClearList("lstAccountName")
    Call ClearList("lstFundYear")
    Call ClearList("lstBankType")
    Call ClearList("lstDesc")
    Call ClearList("lstSch_Grp")
    Call ClearList("lstStatus")
    Call ClearList("lstBDEX")
    Call ClearList("lstCGroup")
    ClearDates (1)
    ClearDates (2)
    ClearDates (3)
End Sub
Private Sub ClearAndRun_Click()
    Call ClearAll_Click
    Call cmdSetFilters_Click
    Call SetListBoxDefaults
End Sub

Private Sub ClearClients_Click()
    Call ClearList("lstAccountName")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearYear_Click()
    Call ClearList("lstFundYear")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearBankType_Click()
    Call ClearList("lstBankType")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearMake_Click()
    Call ClearList("lstDesc")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearSch_Grp_Click()
    Call ClearList("lstSch_Grp")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearBDEX_Click()
    Call ClearList("lstBDEX")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearCGroup_Click()
    Call ClearList("lstCGroup")
    Call cmdSetFilters_Click
End Sub
Private Sub ClearStatus_Click()
    Call ClearList("lstStatus")
    Call cmdSetFilters_Click
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
    SqlA = "SELECT * FROM TFR_Work WHERE (([UNITId] is not null) "
    AccountNameX = "AND ClientId is not null "
    YearX = "AND FYear is not null "
    BankX = "AND Bank is not null "
    StatusX = "AND Assigned_Indicator is not null "
    DescX = "AND UnitDesc is not null "
    Sch_GrpX = "AND Sch_Grp is not null "
    CGroupX = "AND ClientGroupID <> 0 "
    BDEXz = "AND BDEX is not null "

    FundDateRange = "XX"
    BuyOutRangeX = "XX"
    InvoiceSentRangeX = "XX"
    OrderByx = "ORDER BY AccountName"
    CityX = "XX"
    StateX = "XX"
End Sub

Public Sub SETSQL()
    SqlA = "SELECT * FROM TFR_Work WHERE [ClientID] is not null "
    SQLz = SqlA & AccountNameX & YearX & BankX & StatusX & DescX & Sch_GrpX & BDEXz & CGroupX & " "

    If TargetX <> "XX" Then SQLz = SQLz & TargetX
    If FundDateRange <> "XX" Then SQLz = SQLz & FundDateRange
    If BuyOutRangeX <> "XX" Then SQLz = SQLz & BuyOutRangeX
    If BDEXz = "" Then OrderByx = "Order By BDEX" Else OrderByx = "ORDER BY AccountName"
    SQLz = SQLz & OrderByx & ";"
    Dim WhereX As String
    WhereX = Split(SQLz, "TFR_Work ")(1)
    WhereX = Split(WhereX, "Order")(0)

    Call GetSelectedCount("lstAccountName")
    Me.lstAccountName.RowSource = "SELECT AccountName, ClientId From TFR_Work " & WhereX & " GROUP BY AccountName, ClientId ORDER BY AccountName;"
    If Countx > 0 Then Call ReSelectItems("lstAccountName")

    Call GetSelectedCount("lstCGroup")
    Me.lstCGroup.RowSource = "SELECT ClientGroup, ClientGroupID From TFR_Work " & WhereX & " GROUP BY ClientGroup, ClientGroupID ORDER BY ClientGroup;"
    If Countx > 0 Then Call ReSelectItems("lstCGroup")

    Call GetSelectedCount("lstFundYear")
    Me.lstFundYear.RowSource = "SELECT FYear From TFR_Work " & WhereX & " GROUP BY FYear ORDER BY FYear desc;"
    If Countx > 0 Then Call ReSelectItems("lstFundYear")

    Call GetSelectedCount("lstBankType")
    Me.lstBankType.RowSource = "SELECT Bank From TFR_Work " & WhereX & " GROUP BY Bank ORDER BY Bank;"
    If Countx > 0 Then Call ReSelectItems("lstBankType")

    Me.lstStatus.RowSource = "SELECT Assigned_Indicator From TFR_Work " & WhereX & " GROUP BY Assigned_Indicator ORDER BY Assigned_Indicator;"

    Call GetSelectedCount("lstDesc")
    Me.lstDesc.RowSource = "SELECT UnitDesc From TFR_Work " & WhereX & " GROUP BY UnitDesc ORDER BY UnitDesc;"
    If Countx > 0 Then Call ReSelectItems("lstDesc")

    Call GetSelectedCount("lstSch_Grp")
    Me.lstSch_Grp.RowSource = "SELECT Sch_Grp From TFR_Work " & WhereX & " GROUP BY Sch_Grp ORDER BY Sch_Grp;"
    If Countx > 0 Then Call ReSelectItems("lstSch_Grp")

    Call GetSelectedCount("lstBDEX")
    Me.lstBDEX.RowSource = "SELECT BDEX From TFR_Work " & WhereX & " GROUP BY BDEX ORDER BY BDEX;"
    If Countx > 0 Then Call ReSelectItems("lstBDEX")


    Me.Form.RecordSource = SQLz
    Me.UnitCount.Requery
    Me.ReportsList.Requery
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
    For x = 0 To 3
        bName = "Fund" & x
        Me(bName).caption = Right((DatePart("yyyy", Date) + 1) - x, 2)
    Next x
End Sub

Private Sub cmdSetFilters_Click()
Call AccountNameBuild
Call FundYearBuild
Call BankTypeBuild
Call StatusBuild
Call DescBuild
Call BDEXBuild
Call CGroupBuild
 '   Call BuyerBuild
 ' Call TargetPriceSetBuild
    ' Call BuyOutStart_AfterUpdate
 '   Call FundDateStart_AfterUpdate
    ' Call InvoiceSentStart_AfterUpdate
 '   Call CityBuild
 '   Call StateBuild
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
Private Sub CGroupBuild()
    If 1 = 1 Or Me.lstCGroup.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstCGroup
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([ClientGroupID]=" & .Column(1, intRow) & ") OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND (ClientGroupID is not null) "
    End If
    CGroupX = strSQL
End Sub
Private Sub AccountNameBuild()
    If 1 = 1 Or lstAccountName.Column(1) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstAccountName
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "[ClientID]=" & .Column(1, intRow) & " OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([ClientID] is not null) "
    End If
    AccountNameX = strSQL
End Sub
Private Sub FundYearBuild()
    If 1 = 1 Or lstFundYear.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstFundYear
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "[FYear]=" & .Column(0, intRow) & " OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([FYear] > 1) "
    End If
    YearX = strSQL
End Sub
Private Sub BankTypeBuild()
    If 1 = 1 Or Me.lstBankType.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstBankType
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "(Bank='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND (Bank <> '') "
    End If
    BankX = strSQL
End Sub
Private Sub BDEXBuild()
    If 1 = 1 Or Me.lstBDEX.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstBDEX
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "(BDEX='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND (BDEX <> '') "
    End If
    BDEXz = strSQL
End Sub

Private Sub DescBuild()
    If 1 = 1 Or lstDesc.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstDesc
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([UnitDesc]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([UnitDesc] <> '')"
    End If
    DescX = strSQL
End Sub
Private Sub Sch_GrpBuild()
    If 1 = 1 Or lstSch_Grp.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstSch_Grp
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([Sch_Grp]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        'If no results no string else strip off trailing 'and'
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([Sch_Grp] <> '')"
    End If
    Sch_GrpX = strSQL
End Sub

Private Sub StatusBuild()
    If lstStatus.Column(0) & "test" <> "test" Then
        strSQL = " AND ("
        With Me.lstStatus
            For intRow = 0 To .ListCount - 1
                If .Selected(intRow) Then
                    strSQL = strSQL & "([Assigned_Indicator]='" & .Column(0, intRow) & "') OR "
                End If
            Next intRow
        End With
        ' strip off leading comma and space
        If strSQL = " AND (" Then strSQL = "" Else strSQL = Left(strSQL, Len(strSQL) - 4) & ") "
    Else
        strSQL = " AND ([Assigned_Indicator] <> 'x') "
    End If
    StatusX = strSQL
End Sub


Private Sub Fund0_Click()
    Call StartDateRange_Click
    Call SetDateRange("StartDateRange", "EndDateRange", 1)
End Sub
Private Sub Fund1_Click()
    Call StartDateRange_Click
    Call SetDateRange("StartDateRange", "EndDateRange", 0)
End Sub
Private Sub Fund2_Click()
    Call StartDateRange_Click
    Call SetDateRange("StartDateRange", "EndDateRange", -1)
End Sub
Private Sub Fund3_Click()
    Call StartDateRange_Click
    Call SetDateRange("StartDateRange", "EndDateRange", -2)
End Sub
Private Sub SetDateRange(StartField As String, EndField As String, YearBack As Integer)
    Me(StartField) = "1/1/" & DatePart("yyyy", Date) + YearBack
    Me(EndField) = "12/31/" & DatePart("yyyy", Date) + YearBack
    If StartField Like "*Fund*" Then Call StartDateRange_AfterUpdate
End Sub
Private Sub AccountName_DblClick(Cancel As Integer)
    If Me.AccountName & "TEST" = "TEST" Then Exit Sub
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
    Me.UnitID.SetFocus
End Sub

Private Sub Sch_Grp_DblClick(Cancel As Integer)
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "sGrpID=" & Me.GroupID & ""
      Me.UnitID.SetFocus
End Sub
'_________________________
'XLS = [Col letter],[Format Option][Command]
'Col Letter:        Number/letter of the column
'Format Options:    T = Text | B = Bold | $ = Currency | S = Standard Numbers | P = Percent (P# = Percent and the # of decimials)
'Command:           C = Center | L=Left | R=Right | E = SubTotal Col | T= Total Col Count | X = COUNT | ZZ = Delete column | HH = Hide Column


Private Sub SendToExcel_Click()
Dim xSql As String
xSql = Me.Form.RecordSource
xSql = SplitString(xSql, " Order By", 0)
xSql = xSql & " Order By FundingStatus Desc "
'XLS = "|A,ZZ|A,ZZ|A,ZZ|A,ZZ|A,ZZ|A,ZZ|A,ZZ|"
XLS = ""
Call QuickTFR2Excel(xSql)
End Sub

Private Sub StartDateRange_AfterUpdate()
    If IsDate(Me.StartDateRange) And IsDate(Me.EndDateRange) Then
        FundDateRange = "And FundDate Between #" & Me.StartDateRange & "# AND #" & Me.EndDateRange & "# "
        OrderByx = "ORDER BY FundDate"
          Call SETSQL
    Else
        FundDateRange = "XX"
        OrderByx = "ORDER BY AccountName"
    End If
    If Not IsDate(Me.StartDateRange) And Not IsDate(Me.EndDateRange) Then Call SETSQL
End Sub
Private Sub EndDateRange_AfterUpdate()
    StartDateRange_AfterUpdate
End Sub

Private Sub StartDateRange_Click()
    ClearDates (1)
    ClearDates (2)
End Sub
Private Sub EndDateRange_Click()
    ClearDates (1)
    ClearDates (2)
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
        Me.StartDateRange = ""
        Me.EndDateRange = ""
        FundDateRange = "XX"
    End If
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub
