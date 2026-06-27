Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub GetAll_Click()
    For i = 0 To UnitsList.ListCount - 1
        UnitsList.Selected(i) = True
    Next i
End Sub

Private Sub GetAllBack_Click()
    For i = 0 To UnitsFormsList.ListCount - 1
        UnitsFormsList.Selected(i) = True
    Next i
End Sub

Private Sub MoveUnits_Click()
    Dim ItemIndex As Variant
    Call DelWorkTable
    For Each ItemIndex In Me.UnitsList.ItemsSelected
        SQLz = "INSERT INTO MoveUnits_WorkTable ( UnitID ) SELECT " & Me.UnitsList.ItemData(ItemIndex) & " AS UnitID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next
    On Error GoTo fixme1
    SQLz = "UPDATE Units INNER JOIN MoveUnits_WorkTable ON Units.UnitID = MoveUnits_WorkTable.UnitID SET Units.GroupREF = " & Me.NewGroupList & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call RePairExchangeIT
    Me.UnitsFormsList.Requery
    Me.UnitsList.Requery
    Exit Sub
fixme1:
End Sub
Private Sub MoveUnitsBack_Click()
    Dim ItemIndex As Variant
    Call DelWorkTable
    For Each ItemIndex In Me.UnitsFormsList.ItemsSelected
        SQLz = "INSERT INTO MoveUnits_WorkTable ( UnitID ) SELECT " & Me.UnitsFormsList.ItemData(ItemIndex) & " AS UnitID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next
    On Error GoTo fixme1
    SQLz = "UPDATE Units INNER JOIN MoveUnits_WorkTable ON Units.UnitID = MoveUnits_WorkTable.UnitID SET Units.GroupREF = " & Me.GroupList & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call RePairExchangeIT
    Me.UnitsFormsList.Requery
    Me.UnitsList.Requery
    Exit Sub
fixme1:
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From MoveUnits_WorkTable;")
    DoCmd.SetWarnings True

End Sub


Private Sub Form_Load()
    Call ClearList(7)
    Call ClearSearches
End Sub


Private Sub SelectCompany_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig FROM MstrLease WHERE Client=" & Me.SelectCompany & " ORDER By MLNo;"
    Me.LeaseList.RowSource = SQLz
    Call ClearList(6)
    Call ClearSearches
End Sub
Private Sub Leaselist_Click()
    SQLz = "SELECT SchID, SchNo FROM Schedule WHERE SchMLRef=" & Me.LeaseList & " ORDER BY SCHNo1;"
    Me.ScheduleList.RowSource = SQLz
    Me.NewSchedule.RowSource = SQLz
    Call ClearList(5)
    Call ClearSearches
End Sub
Private Sub ScheduleList_Click()
    SQLz = "SELECT SGrpID, SGrpNO FROM SchGrp WHERE SGrpSchID=" & Me.ScheduleList & " ORDER BY SGrpNO ;"
    Me.GroupList.RowSource = SQLz
    Call ClearList(4)
    Call ClearSearches
End Sub
Private Sub GroupList_Click()
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, " & _
    "UnitSubType As SubType FROM Units WHERE GroupREF=" & Me.GroupList & " ORDER BY UnitID;"
    Me.UnitsList.RowSource = SQLz
    Call ClearList(3)
    Call NewSchedule_Click
    Call ClearSearches
End Sub

Private Sub NewSchedule_Click()
    If Me.GroupList = "" Then
        MsgBox "You Must Select the 'Old' Group First", vbInformation, "Select Wisely"
        Exit Sub
    End If
    SQLz = "SELECT SGrpID, SGrpNO FROM SchGrp WHERE SGrpSchID=" & Me.NewSchedule & " AND SGrpID <>" & Me.GroupList & " ORDER BY SGrpNO ;"
    Me.NewGroupList.RowSource = SQLz
    Call ClearList(2)
    Call ClearSearches
End Sub

Private Sub NewGroupList_Click()
    On Error GoTo FixMe
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo FROM Units WHERE GroupREF=" & Me.NewGroupList & " ORDER BY UnitID;"
    Me.UnitsFormsList.RowSource = SQLz
    Call ClearList(1)
    Call ClearSearches
    Exit Sub
FixMe:

End Sub

Private Sub SearchUnitId_Change()
'Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND UnitID like '*" & Me.SearchUnitID.Text & "*' ORDER BY UnitID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchUnitNo_Change()
    Me.SearchUnitID = ""
    'Me.SearchUnitNo = ""
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND UnitUnitNum like '*" & Me.SearchUnitNo.Text & "*' ORDER BY UnitUnitNum, UnitMake;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchVIN_Change()
    Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    'Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""

    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND unitvin like '*" & Me.SearchVin.Text & "*' ORDER BY UnitID, UnitMake;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchYear_Change()
    Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    Me.SearchVin = ""
    'Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""

    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND UnitYr like '*" & Me.SearchYear.Text & "*' ORDER BY UnitYr, UnitMake;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchMake_Change()
    Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    Me.SearchVin = ""
    Me.SearchYear = ""
    'Me.SearchMake = ""
    Me.SearchType = ""

    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND UnitMake like '*" & Me.SearchMake.Text & "*' ORDER BY UnitMake, UnitYr;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchType_Change()
    Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    'Me.SearchType = ""

    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.GroupList & " AND UnitType like '*" & Me.SearchType.Text & "*' ORDER BY UnitType, UnitMake, UnitYr;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchVIN2_Change()
    Me.SearchUnitID2 = ""
    Me.SearchUnitNo2 = ""
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.NewGroupList & " AND unitvin like '*" & Me.SearchVIN2.Text & "*' ORDER BY UnitID;"
    Me.UnitsFormsList.RowSource = SQLz
End Sub
Private Sub SearchUnitId2_Change()
    Me.SearchUnitNo2 = ""
    Me.SearchVIN2 = ""
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.NewGroupList & " AND UnitID like '*" & Me.SearchUnitID2.Text & "*' ORDER BY UnitID;"
    Me.UnitsFormsList.RowSource = SQLz
End Sub
Private Sub SearchUnitNo2_Change()
    Me.SearchUnitID2 = ""
    Me.SearchVIN2 = ""
    SQLz = "SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year, UnitMake As Make, UnitType As Type, UnitSubType As SubType FROM Units " & _
           "WHERE GroupREF=" & Me.NewGroupList & " AND UnitUnitNum like '*" & Me.SearchUnitNo2.Text & "*' ORDER BY UnitID;"
    Me.UnitsFormsList.RowSource = SQLz
End Sub
Private Sub ClearList(Rank As Integer)

    If Rank >= 7 Then Me.SelectCompany = ""

    If Rank >= 6 Then Me.LeaseList = ""
    If Rank >= 6 Then Me.ScheduleList.RowSource = ""
    If Rank >= 6 Then Me.NewSchedule.RowSource = ""

    If Rank >= 5 Then Me.ScheduleList = ""
    If Rank >= 5 Then Me.NewSchedule = ""
    If Rank >= 5 Then Me.GroupList.RowSource = ""
    If Rank >= 5 Then Me.NewGroupList.RowSource = ""

    If Rank >= 4 Then Me.UnitsList.RowSource = ""
    If Rank >= 4 Then Me.GroupList = ""

    If Rank >= 3 Then Me.UnitsList = ""

    If Rank >= 2 Then Me.UnitsFormsList.RowSource = ""
    If Rank >= 2 Then Me.NewGroupList = ""
    If Rank >= 1 Then Me.UnitsFormsList = ""


End Sub
Private Sub ClearSearches()
    Me.SearchUnitID = ""
    Me.SearchUnitNo = ""
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchUnitID2 = ""
    Me.SearchUnitNo2 = ""
    Me.SearchVIN2 = ""
End Sub
