Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, TargetQuery, FormX As String
Dim RST As DAO.Recordset

Public Sub SetPerDiemsTypeInSch()
    If Me.SchID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "MassMod_PerdiemType"
    DoCmd.OpenForm "MassMod_PerdiemType", , , , , , Me.SchID
End Sub
Private Sub ClientList_DblClick(Cancel As Integer)
    Call ClientView_Click
End Sub

Public Sub ClientView_Click()
    If Me.ClientList.Column(0) & "test" = "test" Then Exit Sub
    If Me.UnitSelection < 3 Then FormX = "ClientsView" Else FormX = "ClientsView"
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientList.Column(0)
End Sub

Private Sub cmdOpenGroupNew_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "GroupForm2"
    DoCmd.OpenForm "GroupForm2", , , "SGrpID=" & Me.GroupList.Column(0)
End Sub

Private Sub CloneSchGrp_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    If FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Or FAUserAccessGroup Like "*OPS*" Then
        AddReportLog "CloneSchGrp_Click"
        DoCmd.close acForm, "CloneScheduleGroups"
        DoCmd.OpenForm "CloneScheduleGroups", , , , , , Me.SchID & "|" & Me.groupID
    End If
End Sub

Private Sub EditPerDiems_Click()
 If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
 DoCmd.close acForm, "MassMod_PerdiemRents"
 DoCmd.OpenForm "MassMod_PerdiemRents", , , , , , Me.groupID & "?" & Me.SchNo
End Sub

Private Sub ExitOptionsDashboard_Click()
    Dim OpenArgX
    OpenArgX = "FALease"
    If Me.LeaseList.Column(2) <> "FA" Then OpenArgX = "Portfolio"
    DoCmd.OpenForm "ExitOptionsManager", , , , , , "OpenArgX"
    If Me.ClientID & "test" <> "test" Then
        AddReportLog "ExitOptionsDashboard_Click"
        Forms!ExitOptionsManager.ClientList = Me.ClientID
        Form_ExitOptionsManager.ClientList_Click
        If OpenArgX = "Portfolio" Then Forms!ExitOptionsManager.LeaseTypeToShow = 2 Else Forms!ExitOptionsManager.LeaseTypeToShow = 1
        Form_ExitOptionsManager.LeaseTypeToShow_Click
        If Me.groupID & "test" <> "test" Then
            Forms!ExitOptionsManager.SchGrpList = Me.groupID
            Form_ExitOptionsManager.SchGrpList_Click
        End If
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.SchNo = ""
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub GroupList_DblClick(Cancel As Integer)
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.groupID, , , Me.groupID
End Sub

Private Sub LeaseList_DblClick(Cancel As Integer)
    If Me.LeaseList & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "MstrLease"
    DoCmd.OpenForm "MstrLease", , , "MLID=" & Me.LeaseList.Column(0)
End Sub

Public Sub View_Active_ATF_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    Dim FAIDx As String
    FAIDx = Nz(DLookup("FAID", "vw_SixKeys", "MLOrig= 'FA' and Schid=" & Me.SchID), "")
    If FAIDx = "" Then
        MsgBox "This NOT an FA Leased Schedule", vbCritical, AtlaasVersion
        Exit Sub
    Else
        AddReportLog "View_Active_ATF_Click"
        Call OpenActiveATFReport(Me.SchID)
    End If
End Sub

Private Sub MoveUnitsForm_Click()
    AddReportLog "MoveUnitsForm_Click"
    DoCmd.OpenForm "MoveUnit"
End Sub

Private Sub OpenMassMod_Click()
    If Me.groupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    AddReportLog "OpenMassMod_Click"
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = Me.ClientID
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = Me.MLID
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = Me.SchID
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = Me.GroupList
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click

End Sub
Private Sub ReportsList_Click()
    Me.ReportsList.ControlTipText = Nz(ReportsList.Column(1), "test")
End Sub

Private Sub ReportsList_DblClick(Cancel As Integer)
    If Me.ScheduleList & "test" = "test" Then Exit Sub
    Dim FormX, CommandX As String
    FormX = Me.ReportsList.Column(3)
    CommandX = Me.ReportsList.Column(2)
    If FormX = "Local" Then
        CallByName Me, CommandX, VbMethod
    Else
        CallByName FormX, CommandX, VbMethod
    End If
    AddReportLog CommandX
End Sub
Public Sub ExcelUnitPartInterimReport()
    If Me.SchID & "test" = "test" Then Exit Sub
    Call MakeDrawDownUnitReport("SchID", CLng(Me.SchID), "Unit Draw-Down Report - for Client:" & DLookup("ClientShortName", "vw_SixKeys", "SchId=" & Me.SchID) & "Sch: " & DLookup("Schedule", "vw_SixKeys", "SchId=" & Me.SchID))
End Sub
Public Sub ClientUnitSummary_Click()
Dim XLS As String
    Dim Countx As Integer
SQLz = "Select * From Client_LPAF Where ClientId=" & Me.ClientID & ";"
Countx = DCount("clientID", "Client_LPAF", "ClientId=" & Me.ClientID)
If Nz(Countx, 0) <= 0 Then
MsgBox "No Matching Units", vbInformation, "No Result"
Exit Sub
End If
XLS = "|A,BC|B,BC|G,$E|H,$E|J,BC|M,$E|N,BC|O,$E|P,BC|Q,$E|R,$E|s,HH|T,HH|U,HH|"
AddReportLog "ClientUnitSummary_Click"
SendReport2Excel SQLz, "Report", "Report for: " & Me.ClientList.Column(1) & " All Schedules - Units: " & Countx, True, XLS, Countx
End Sub
Public Sub MakeRALetter()
    If Me.SchID & "test" = "test" Then Exit Sub
    If Me.LeaseList.Column(2) <> "FA" Then
        MsgBox "A Rent Assigment Letter can only be ran on an 'FA' Lease", vbCritical, AtlaasVersion
        Exit Sub
    End If
    AddReportLog "MakeRALetter"
    Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Run the Rent Assigment Letter</font></b><br/>"
    msg1 = msg1 & "Which 'RAL' do You want to run?</b><br /><br/>"
    msg1 = msg1 & "<ul><b>RAL?</b><i> RAL with Word Unit Attachment</i></ul>"
    msg1 = msg1 & "<ul><b>RAL w/EXCEL?</b><i> RAL w/Excel Spreadsheet</i></ul>"
    msg1 = msg1 & "<ul><b>Just EXCEL?</b><i> Just the Excel Spreadhseet</i></ul><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "RAL", "RAL w/Excel", "Just Excel")
    If ResposeX = vbBt1 Then ' RAL Letter w/Attach
        Call MakeRAL
        Call MakeRALAttachment
    End If
    If ResposeX = vbBt2 Then ' RAL Letter w/Excel
        Call MakeRAL
        Call MakeRALAttachmentEXCEL(Me.SchID)
    End If
    If ResposeX = vbBt3 Then  ' Just Excel
        Call MakeRALAttachmentEXCEL(Me.SchID)
    End If
End Sub
Private Sub ScheduleList_DblClick(Cancel As Integer)
    If Me.SchID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & Me.SchID
End Sub

Public Sub RunScheduleSyndCheck()
    If Me.SchID & 0 = 0 Then Exit Sub
    AddReportLog "RunScheduleSyndCheck"
    Call SynCheckExcel(Me.SchID)
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count],ClientGroupID " & _
           "From " & TargetQuery & " GROUP BY clientID, ClientShortName, CompanyName, ClientGroupID " & _
           "Having CompanyName Like '*" & Me.SearchClient.Text & "*' Or ClientShortName Like '*" & Me.SearchClient.Text & "*' ORDER BY ClientShortName;"
    Me.ClientList.RowSource = SQLz
End Sub


Private Sub SignOff_Click()
'Exit Sub
    If Me.groupID & 0 = 0 Then
        MsgBox "You must select a Group before you a sign off on it.", vbCritical, AtlaasVersion
        Exit Sub
    End If
    If DLookup("MLOrig", "vw_FiveKeys", "GroupID=" & Me.groupID & "") = "FA" Then
    AddReportLog "SignOff_Click"
        If Right(Me.SignOffStatus, 1) = "0" Then
            Call StartSignOff(Me.groupID, "SignOff1")
        Else
            Call StartSignOff(Me.groupID, "SignOff" & Right(SignOffStatus, 1))
        End If
    End If
End Sub
Public Sub LastTitleEvent_Click()
  If Me.SchID & 0 = 0 Then Exit Sub
  AddReportLog "LastTitleEvent_Click"
  Call LastTitleEventReport(Me.SchID)
End Sub


Private Sub ToggleSortAssetID1_Click()
    Call SortListBox("ToggleSortAssetID1", "Units_List", "AssetID")
End Sub
Private Sub ToggleSortVIN1_Click()
    Call SortListBox("ToggleSortVIN1", "Units_List", "VIN")
End Sub
Private Sub ToggleSortUnitNum1_Click()
    Call SortListBox("ToggleSortUnitNum1", "Units_List", "UnitNum")
End Sub
Private Sub ToggleSortYear1_Click()
    Call SortListBox("ToggleSortYear1", "Units_List", "MYear")
End Sub
Private Sub ToggleSortAssetID2_Click()
    Call SortListBox("ToggleSortAssetID2", "UnitLookUp", "AssetID")
End Sub
Private Sub ToggleSortVIN2_Click()
    Call SortListBox("ToggleSortVIN2", "UnitLookUp", "VIN")
End Sub
Private Sub ToggleSortUnitNum2_Click()
    Call SortListBox("ToggleSortUnitNum2", "UnitLookUp", "UnitNum")
End Sub
Private Sub ToggleSortPlate2_Click()
    Call SortListBox("ToggleSortPlate2", "UnitLookUp", "UnitPlateNum")
End Sub
Private Sub ToggleSortClient2_Click()
    Call SortListBox("ToggleSortClient2", "UnitLookUp", "ClientShortName")
End Sub

Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Public Sub ClearToggles(WhichToggles As Integer)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" Then
            If WhichToggles = 1 And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If WhichToggles = 2 And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
            If WhichToggles = 0 Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
End Sub
Private Sub UnitLookUp_DblClick(Cancel As Integer)
    If Me.UnitLookUp.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitLookUp.Column(0))
    If Me.UnitSelection > 2 Then Forms!Unitsform.UnitsTabControl.Value = 6
    Me.SearchVin.SetFocus
End Sub

Private Sub Units_List_Click()
    If Me.Units_List.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    SQLz = "SELECT * From Units WHERE UnitID=" & Me.Units_List.Column(0) & ";"
    Me.ScheduleUnitInfo_SubForm.Form.RecordSource = SQLz
    Me.ScheduleUnitInfo_SubForm.Visible = True
End Sub

Private Sub ViewMasterLease_Click()
    If Me.LeaseList.Column(0) & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "MstrLease"
    DoCmd.OpenForm "MstrLease", , , "MLID=" & Me.LeaseList.Column(0)
End Sub

Private Sub ViewSchedules_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & Me.SchID
End Sub

Private Sub ViewUnitGroups_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.groupID, , , Me.SchID
End Sub


Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()

'Call DelWorkTable
    Me.ScheduleList.RowSource = ""
    Me.LeaseList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SchNo = ""
    Me.SearchSchedule = ""
    Me.UnitLookUp.RowSource = ""
    Me.SignOffForm.Visible = False
    TargetQuery = "vw_SixKeys"
    Call ClearToggles(0)
    Call SetClientList
    Call HideShow
End Sub

Private Sub SetClientList()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
           "FROM " & TargetQuery & " " & _
           "GROUP BY clientID, ClientShortName, CompanyName " & _
           "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
    Call DelWorkTable
    Me.ScheduleList.RowSource = ""
    Me.LeaseList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SearchSchedule = ""
End Sub

Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From ReMarketingBuyOut_WorkTable;")
    DoCmd.SetWarnings True
End Sub


Public Sub SchCheck_Click()
    Dim ReportX As String
    If Me.SchID & "test" = "test" Then Exit Sub
    If Me.ClientID = 4029 Then ReportX = "SchedCheckTLS" Else ReportX = "SchedCheck2"
    AddReportLog "SchCheck_Click"
    DoCmd.OpenReport ReportX, acViewPreview, "", "SchID=" & Me.SchID
End Sub

Public Sub ScheduleCheckLists_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    Call MakeWorkTable
    DoCmd.OpenReport "ScheduleCheckListNew", acViewPreview, , "SchID=" & Me.SchID, acNormal
    AddReportLog "ScheduleCheckLists_Click"
End Sub

Public Sub SerialNumbers_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    AddReportLog "SerialNumbers_Click"
    DoCmd.OpenReport "ScheduleSerialNumbers", acViewPreview, "", "[SchID]=" & Me.SchID, acNormal
End Sub
Public Sub Custodian_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Run Custodian</font></b><br/>"
    msg1 = msg1 & "Which Custodian Report do You want to Run?</b><br /><br/>"
    msg1 = msg1 & "  <b>Lablels or Regular?</b><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Lablels", "Regular")
    If ResposeX = vbBt2 Then
        Call MakeWorkTable
        AddReportLog "Custodian_Click"
        Call MakeCustodianFiles(Me.SchID)
    Else
        DoCmd.close acReport, "LabelsToCustodian"
        DoCmd.OpenReport "LabelsToCustodian", acViewPreview
    End If
End Sub


Private Sub Cost_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    AddReportLog "Cost_Click"
    FormatExcel "ScheduleCost", "C:\Faas\" & FAUserName & "\Temp\ScheduleCost.xlsx", "ScheduleCost", , , , "Schedule Unit Cost", True
End Sub
Public Function MakeDocsReport(Optional TwoTab As Boolean)
    If Me.SchID & "test" = "test" Then Exit Function
    Call MakeWorkTable
    AddReportLog "MakeDocsReport"
    If Me.ReportsList.Column(1) = "Docs Report 2 tab" Then
        Call MakeScheduleReport(Me.SchID, True)
    Else
        Call MakeScheduleReport(Me.SchID)
    End If
End Function
Public Sub OpsUnitList_Click()
    AddReportLog "OpsUnitList_Click"
    Call MakeWorkTable
    If Me.SchID & "test" = "test" Then Exit Sub
        Dim ResposeX As VbMsgBoxResultEx
    Dim msg1   As String
    msg1 = msg1 & "<b><Font Size=3>Run OPS Report</font></b><br/>"
    msg1 = msg1 & "Which OPS Report do You want to Run?</b><br /><br/>"
    msg1 = msg1 & "  <b>Long or Short?</b><br/>"
    msg1 = msg1 & "  Please select..."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Long", "Short")
    If ResposeX = vbBt2 Then Call MakeOPSReport(, Me.SchID, True) ' Short
    If ResposeX = vbBt1 Then Call MakeOPSReport(, Me.SchID)  ' Long
End Sub
Private Sub MakeWorkTable()
    DoCmd.SetWarnings False
    SQLz = "Delete * From Schedule_Group_WorkTable;"
    DoCmd.RunSQL (SQLz)
    SQLz = "INSERT INTO Schedule_Group_WorkTable ( ScheduleId ) SELECT " & Me.SchID & " AS SchId;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub


Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID, ClientGroupId From " & TargetQuery & " " & _
           "GROUP BY MLID, MLNo, MLOrig, ClientGroupId, ClientId Having ClientId =" & Me.ClientList & " ORDER BY MLOrig, MLNo;"
    Me.LeaseList.RowSource = SQLz
    Me.ScheduleList.RowSource = ""
    Me.SignOffForm.Form.RecordSource = "Select * From vw_UnitGroupHealth Where GroupId=0"
    Me.ScheduleInfo_SubForm.Form.RecordSource = "Select * From ScheduleInfo Where SchID =0;"
    Me.ClientID = Me.ClientList
    Me.ClientGroupID = Me.ClientList.Column(4)
    Me.MLID = ""
    Me.SchID = ""
    Me.groupID = ""
    Me.Units_List.RowSource = ""
    Me.SignOffStatus = ""
    Me.SearchSchedule = ""
    If Not IsNull(Me.ClientList) Then Me.SchNo = "<b>" & Me.ClientList.Column(1) & "</b>"
    Call ClearToggles(0)
End Sub

Public Sub Leaselist_Click()
    On Error GoTo FixMe
    SQLz = "SELECT SchID, Schedule, Count(AssetID) AS Units, MLID " & _
           "From " & TargetQuery & " " & _
           "GROUP BY SchID, Schedule, MLID " & _
           "HAVING MLID=" & Me.LeaseList & " Order BY CDbl(NumInString(Left(Schedule,3))) desc, Schedule desc; "
    Me.ScheduleList.RowSource = SQLz
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SignOffForm.Form.RecordSource = "Select * From vw_UnitGroupHealth Where GroupId=0"
    Me.SignOffStatus = ""
    Me.ScheduleInfo_SubForm.Form.Visible = False
    Me.ClientGroupID = Me.LeaseList.Column(4)
    Me.MLID = Me.LeaseList.Column(0)
    Me.SearchSchedule = ""
    Me.SchID = ""
    Me.groupID = ""
    Me.SchNo = "<b>" & Me.ClientList.Column(1) & "</b>" & " ML:<b> " & Me.LeaseList.Column(1) & "</b>"
    Call HideShow
    Me.TotalUnits = DCount("VIN", "vw_SixKeys", "MLID=" & Me.LeaseList & "")
    Call ClearToggles(1)
FixMe:
End Sub
Public Sub HideShow()
    If Me.LeaseList.Column(2) = "FA" Then Me.SignOffForm.Visible = True Else Me.SignOffForm.Visible = False
    If FAUserAccessLevel > 11 Then
        Me.MoveUnitsForm.Visible = True
        Me.OpenMassMod.Visible = True
        If FAUserAccessLevel > 90 Then Me.ExitOptionsDashboard.Visible = True
    Else
        Me.MoveUnitsForm.Visible = False
        Me.ExitOptionsDashboard.Visible = False
    End If
    Me.ScheduleUnitInfo_SubForm.Visible = False
End Sub

Public Sub ScheduleList_Click()
    SQLz = "SELECT SGrpID, UnitGroup As [Group], Count(AssetID) AS Units, SchID, BLCD " & _
           "From " & TargetQuery & " " & _
           "GROUP BY SGrpID, UnitGroup, SchID, BLCD " & _
           "HAVING SchID=" & Me.ScheduleList & " Order BY UnitGroup; "
    Me.GroupList.RowSource = SQLz
    Me.groupID = ""
    Me.SchNo = "<b>" & Me.ClientList.Column(1) & "</b>" & " ML:<b> " & Me.LeaseList.Column(1) & "</b> Sch:<b> " & Me.ScheduleList.Column(1) & "</b>"
    Me.Units_List.RowSource = ""
    Me.SchID = Me.ScheduleList
    Me.Units_List.RowSource = ""
    Me.ScheduleInfo_SubForm.Form.Visible = True
    SQLz = "Select * From ScheduleInfo Where SchID =" & Me.ScheduleList & ";"
    Me.ScheduleInfo_SubForm.Form.RecordSource = SQLz
    Me.SignOffStatus = ""
    Me.SignOffForm.Form.RecordSource = "Select * From vw_UnitGroupHealth Where GroupId=0"
End Sub
Private Sub GroupList_Click()
    SQLz = "SELECT AssetID, VIN, UnitNum,  Make, MYear, SGrpID " & _
           "From " & TargetQuery & " " & _
           "GROUP BY AssetID, VIN, UnitNum, Make, MYear, SGrpID " & _
           "HAVING SGrpID=" & Me.GroupList & " Order By AssetID;"
    Me.Units_List.RowSource = SQLz
    Me.groupID = Me.GroupList.Column(0)
    Me.SignOffStatus = "Sign Off Level " & Nz(Right(DMax("Department", "GroupFinalizedTracker", "GroupID=" & Me.groupID & ""), 1), 0)
    Me.SchNo = "<b>" & Me.ClientList.Column(1) & "</b>" & " ML:<b> " & Me.LeaseList.Column(1) & "</b> Sch:<b> " & Me.ScheduleList.Column(1) & "</b> Grp:<b> " & Me.GroupList.Column(1) & "</b>"
    Me.SignOffForm.Form.RecordSource = "Select * From vw_UnitGroupHealth Where GroupId=" & Me.groupID
    Call HideShow
End Sub

Private Sub ClearLabel_Click()
    Me.SearchAssetID = ""
    Me.SearchVin = ""
    Me.SearchUnitNum = ""
    Me.SearchMYear = ""
    Me.SearchSchedule = ""
    Me.SearchVin.SetFocus
    Call ClearToggles(1)
    Call SearchUnits_List("VIN")
End Sub
Private Sub ClearLabel2_Click()
    Me.SearchAssetID2 = ""
    Me.SearchVIN2 = ""
    Me.SearchUnitNum2 = ""
    Me.SearchUnitPlateNum2 = ""
    Me.SearchVIN2.SetFocus
    Me.UnitLookUp.RowSource = ""
    Call ClearToggles(2)
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
    If Me.Units_List.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.Units_List.Column(0))
    If Me.UnitSelection > 2 Then Forms!Unitsform.UnitsTabControl.Value = 6
    Me.SearchVin.SetFocus
End Sub
Public Function SearchUnits_List(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetID = ""
    If sField <> "VIN" Then Me.SearchVin = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum = ""
    If sField <> "MYear" Then Me.SearchMYear = ""
    If Len(Me("Search" & sField).Text) >= 1 And Len(Me("Search" & sField).Text) < 3 Then Exit Function
    SQLz = "SELECT AssetID, VIN, UnitNum, Make, MYear, SchID " & _
           "FROM vw_SixKeys " & _
           "Where sGrpID=" & Me.groupID & " AND " & sField & " Like '*" & Me("Search" & sField).Text & "*' Order BY " & sField & ";"
    Me.Units_List.RowSource = SQLz
End Function
Public Function SearchUnitsLookUp(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetID2 = ""
    If sField <> "VIN" Then Me.SearchVIN2 = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum2 = ""
    If sField <> "UnitPlateNum" Then Me.SearchUnitPlateNum2 = ""
    If Len(Me("Search" & sField & "2").Text) >= 1 And Len(Me("Search" & sField & "2").Text) < 3 Then Exit Function
    If Len(Me("Search" & sField & "2").Text) = 0 Then
       Me.UnitLookUp.RowSource = ""
       Exit Function
    End If
    SQLz = "SELECT AssetID, VIN, UnitNum, UnitPlateNum, ClientShortName, clientID, SchID, SGrpID, MLID " & _
           "FROM vw_SixKeys " & _
           "Where " & sField & " Like '*" & Me("Search" & sField & "2").Text & "*' Order BY " & sField & ";"
    Me.UnitLookUp.RowSource = SQLz
End Function
Private Sub SearchAssetId_Change()
    Call SearchUnits_List("AssetID")
End Sub
Private Sub SearchVIN_Change()
    Call SearchUnits_List("VIN")
End Sub
Private Sub SearchUnitNum_Change()
    Call SearchUnits_List("UnitNum")
End Sub
Private Sub SearchMYear_Change()
    Call SearchUnits_List("MYear")
End Sub
Private Sub SearchSchedule_Change()
On Error GoTo FixMe
    SQLz = "SELECT SchID, Schedule, Count(AssetID) AS Units, MLID " & _
           "From vw_SixKeys " & _
           "WHERE Schedule Like '*" & Me.SearchSchedule.Text & "*' " & _
           "GROUP BY SchID, Schedule, MLID " & _
           "HAVING MLID=" & Me.LeaseList & " Order BY CDbl(NumInString(Left(Schedule,3))) desc, Schedule desc; "
    Me.ScheduleList.RowSource = SQLz
FixMe:
End Sub
Private Sub SearchAssetID2_Change()
    Call SearchUnitsLookUp("AssetID")
End Sub
Private Sub SearchVIN2_Change()
    Call SearchUnitsLookUp("VIN")
End Sub
Private Sub SearchUnitNum2_Change()
    Call SearchUnitsLookUp("UnitNum")
End Sub
Private Sub SearchUnitPlateNum2_Change()
    Call SearchUnitsLookUp("UnitPlateNum")
End Sub
Public Sub Remotelookup()
    Call UnitLookUp_Click
End Sub
Private Sub UnitLookUp_Click()
    If Me.UnitLookUp.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If

    If Me.UnitSelection <> 1 Then Me.UnitSelection = 1
    Call UnitSelection_Click
    Me.ClientList = Me.UnitLookUp.Column(5)
    Call ClientList_Click
    Me.LeaseList = Me.UnitLookUp.Column(8)
    Call Leaselist_Click
    Me.ScheduleList = Me.UnitLookUp.Column(6)
    Call ScheduleList_Click
    Me.GroupList = Me.UnitLookUp.Column(7)
    Call GroupList_Click

    For i = 0 To Units_List.ListCount - 1
        If (Me.Units_List.Column(0, i)) = Me.UnitLookUp.Column(0) Then
            Me.Units_List.Selected(i) = True
            Me.Units_List.SetFocus
            Me.Units_List.ListIndex = (i)
        End If
    Next i
    Call Units_List_Click
End Sub
Public Sub UnitSelection_Click()
    If Me.UnitSelection = 1 Then TargetQuery = "vw_SixKeys"
    If Me.UnitSelection = 2 Then TargetQuery = "vw_Sch_Grp_Units_Leased"
    If Me.UnitSelection = 3 Then TargetQuery = "vw_SixKeys_bUnits"
    If Me.UnitSelection = 4 Then TargetQuery = "vw_Sch_Grp_Units_Tracking"
    If Me.UnitSelection = 5 Then TargetQuery = "vw_Sch_Grp_Units_TBReturned"
    Call SetClientList
    Call ClientList_Click
End Sub
