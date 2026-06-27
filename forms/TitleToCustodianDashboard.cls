Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim RST        As DAO.Recordset
Dim SQLi, SQLx, SQLy, SQLz As String
Dim CGrpInt As Long


Private Sub AssetID_DblClick(Cancel As Integer)
Call OpenUnitsForm
End Sub

Private Sub ClearFilters_Click()
    Call Form_Load
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub SortBDEX1_Click()
Call SortForm("SortBDEX1", "BDEX")
End Sub

Private Sub TitlesOutReport_Click()
    Dim SQLz, XLS, Headerx As String
    Call MessageUser("Open", "Getting Data for Export", "Just a moment...")
    Headerx = "Title Inventory IN / OUT - "
    If Me.SearchClient <> "" Then Headerx = Headerx & "Client: " & Me.SearchClient.Column(1) & " "
    If Me.SearchSchNo <> "" Then Headerx = Headerx & "|Sch: " & Me.SearchSchNo.Column(1) & " "
    If Me.SearchGroup <> "" Then Headerx = Headerx & "|Grp: " & Me.SearchGroup.Column(1) & " "
    If Me.SearchEvent <> "" Then Headerx = Headerx & "|Event: " & Me.SearchEvent.Column(0) & " "
    Headerx = Headerx & "as of : " & Format(Date, "mm-dd-yy")
    XLS = "|A,DD|F,BC|G,BC|H,BC|H,cColor:3:204.102.0|I,SR|"
    XLS = XLS & "G,cLabel:End=G:3:IN From" & vbNewLine & "State|"
    XLS = XLS & "H,cLabel:End=H:3:Appl" & vbNewLine & "OUT|"
    XLS = XLS & "I,cLabel:End=I:3:Days" & vbNewLine & "Elapsed|"
    XLS = XLS & "J,cLabel:End=J:3:First" & vbNewLine & "IN|"
    XLS = XLS & "K,cLabel:End=K:3:Last" & vbNewLine & "IN|"
    XLS = XLS & "L,cLabel:End=L:3:Fund" & vbNewLine & "Status|L,BC|"
    XLS = XLS & "N,cLabel:End=N:3:Funded" & vbNewLine & "Date|O,ZZ|O,ZZ|O,ZZ|O,ZZ|O,ZZ|"
    SQLz = "Select * FROM vw_TitleToCustodian_Pending "
    SQLz = SQLz & "WHERE " & SplitString(Me.Form.RecordSource, "WHERE", 1)
    SQLz = Replace(SQLz, "vw_TitleToCustodian_Pending", "vw_TitleToCustodian_PendingINOUT")
    SendReport2Excel SQLz, "Titled", Headerx, True, XLS, 0, True
    Call MessageUser("Close")
End Sub
Private Sub ExcelLegal_Click()
Call SendExcel
End Sub
Private Function SendExcel()
    Dim SQLz, XLS, Headerx As String
SQLz = "Select "
SQLz = SQLz & " Client"
SQLz = SQLz & ",Schedule"
SQLz = SQLz & ",UnitGroup"
SQLz = SQLz & ",AssetID"
SQLz = SQLz & ",UnitStatus"
SQLz = SQLz & ",UnitDesc"
SQLz = SQLz & ",LastEvent"
SQLz = SQLz & ",TitleLog"
SQLz = SQLz & ",DaysElapsed"
SQLz = SQLz & ",VIN"
SQLz = SQLz & ",Assignee"
SQLz = SQLz & ",FundedDate"
SQLz = SQLz & " FROM vw_TitleToCustodian_Pending "
SQLz = SQLz & "WHERE " & SplitString(Me.Form.RecordSource, "WHERE", 1)
    Call MessageUser("Open", "Setting Up Export", "Wait, wait, wait for it...")
    Headerx = "Title Inventory Pending - "
    If Me.SearchClient <> "" Then Headerx = Headerx & "Client: " & Me.SearchClient.Column(1) & " "
    If Me.SearchSchNo <> "" Then Headerx = Headerx & "|Sch: " & Me.SearchSchNo.Column(1) & " "
    If Me.SearchGroup <> "" Then Headerx = Headerx & "|Grp: " & Me.SearchGroup.Column(1) & " "
    If Me.SearchEvent <> "" Then Headerx = Headerx & "|Event: " & Me.SearchEvent.Column(0) & " "
    Headerx = Headerx & "as of : " & Format(Date, "mm-dd-yy")
    XLS = "|A,DD|D,TX|D,TR|E,BC|I,BC|"
    XLS = XLS & "C,cLabel:End=C:3:Unit" & vbNewLine & "Group|"
    XLS = XLS & "E,cLabel:End=E:3:Unit" & vbNewLine & "Status|"
    XLS = XLS & "G,cLabel:End=G:3:Last Title" & vbNewLine & "Event|"
    XLS = XLS & "H,cLabel:End=H:3:Last Event" & vbNewLine & "Date|"
    XLS = XLS & "I,cLabel:End=I:3:Days" & vbNewLine & "Elapsed|"
    XLS = XLS & "L,cLabel:End=L:3:Funded" & vbNewLine & "Date|"
    'XLS = XLS & "A,Banding|"
    SendReport2Excel SQLz, "Titled", Headerx, True, XLS, 0, True
    Call MessageUser("Close")
End Function

Private Sub VIN_DblClick(Cancel As Integer)
Call OpenUnitsForm
End Sub

Private Sub Form_Current()
    Me.txtCurrent = Me.AssetID
End Sub

Public Sub Form_Load()
    Me.SearchSchNo.RowSource = ""
    Me.SearchGroup.RowSource = ""
    Me.Form.RecordSource = "Select * FROM vw_TitleToCustodian_Pending WHERE LastEvent<>'' Order By Client"
    Me.SearchEvent.RowSource = "Select LastEvent, Count(AssetID) FROM vw_TitleToCustodian_Pending WHERE LastEvent<>'' Group By LastEvent Order By LastEvent"
    Me.SearchClient = ""
    Me.SearchSchNo = ""
    Me.SearchGroup = ""
    Me.SearchEvent = ""
ClearSorts (1)
End Sub

Private Sub FundedDate_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub

Private Sub Assignee_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub

Private Sub Lessee_DblClick(Cancel As Integer)
Call OpenClientsForm
End Sub

Private Sub SchNo_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub
Private Sub Client_DblClick(Cancel As Integer)
Call OpenClientsForm
End Sub

Private Sub Status_DblClick(Cancel As Integer)
Call OpenUnitsForm
End Sub

Private Sub UnitDesc_DblClick(Cancel As Integer)
Call OpenUnitsForm
End Sub

Private Sub UnitGroup_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub OpenUnitsForm()
'DoCmd.OpenForm "UnitsForm", acNormal, , "UnitID = " & Me.AssetID
End Sub
Private Sub OpenScheduleForm()
DoCmd.close acForm, "Schedule_Form"
DoCmd.OpenForm "Schedule_Form", , , "Schid=" & Me.SchID & ""
Forms!Schedule_Form.GroupFundList.SetFocus
End Sub
Private Sub OpenGroupForm()
DoCmd.close acForm, "GroupForm"
DoCmd.OpenForm "GroupForm", , , "Schid=" & Me.SchID & ""
End Sub
Private Sub OpenClientsForm()
DoCmd.close acForm, "ClientsView"
DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
End Sub


Private Function OpenMassMod(GroupIDx As Long)
    If Me.SchID & "test" = "test" Then
        MsgBox "Sorry, but you must first select row", , "No Selections, No Service :("
        Exit Function
    End If
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = Me.clientID
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = DLookup("MLID", "vw_Sixkeys", "sGrpID=" & GroupIDx)
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = Me.SchID
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = GroupIDx
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
End Function



Private Sub SearchClient_Click()
    Me.SearchSchNo = ""
    Me.SearchGroup = ""
    ClearSorts (1)
    Call SetFormData("Client")

    SQLz = "Select SchID, Schedule From vw_TitleToCustodian_Pending"
    SQLz = SQLz & " Where ClientID=" & Me.SearchClient
    SQLz = SQLz & " Group By SchID, Schedule Order By Schedule"
    Me.SearchSchNo.RowSource = SQLz

    SQLz = "Select SGrpID, UnitGroup From vw_TitleToCustodian_Pending"
    SQLz = SQLz & " Where ClientID=" & Me.SearchClient
    SQLz = SQLz & " Group By SGrpID, UnitGroup Order By UnitGroup"
    Me.SearchGroup.RowSource = SQLz
End Sub
Public Function SetFormData(Typex As String)
    SQLz = "Select * FROM vw_TitleToCustodian_Pending"
    SQLz = SQLz & " WHERE ClientID >0"
    If SearchClient <> "" Then SQLz = SQLz & " AND clientID = " & SearchClient
    If SearchSchNo <> "" Then SQLz = SQLz & " AND SchID=" & Me.SearchSchNo.Column(0)
    If SearchGroup <> "" Then SQLz = SQLz & " AND sGrpID =" & SearchGroup.Column(0)
    If SearchEvent <> "" Then SQLz = SQLz & " AND LastEvent = '" & SearchEvent.Column(0) & "' "
    If Typex = "Client" Then SQLz = SQLz & " Order By Schedule Desc"
    If Typex = "Schedule" Then SQLz = SQLz & " Order By Schedule, UnitGroup"
    If Typex = "Group" Then SQLz = SQLz & " Order By UnitGroup"
    If Typex = "Event" Then SQLz = SQLz & " Order By LastEvent"
    Me.Form.RecordSource = SQLz
End Function
Private Sub SearchSchNo_Click()
    Me.SearchGroup = ""
    ClearSorts (1)
    Call SetFormData("Schedule")
    SQLz = "Select SGrpID, UnitGroup From vw_TitleToCustodian_Pending"
    SQLz = SQLz & " Where ClientID=" & Me.SearchClient
    SQLz = SQLz & " Group By SGrpID, UnitGroup Order By UnitGroup"
    Me.SearchGroup.RowSource = SQLz
End Sub
Private Sub SearchGroup_Click()
    ClearSorts (1)
    Call SetFormData("Group")
    SQLz = "Select * FROM vw_TitleToCustodian_Pending"
    SQLz = SQLz & " WHERE sGrpID =" & SearchGroup.Column(0)
    SQLz = SQLz & " Order By UnitGroup"
    Me.Form.RecordSource = SQLz
End Sub
Private Sub SearchEvent_Click()
    ClearSorts (1)
    Call SetFormData("Event")
End Sub

Private Sub SortAssignee1_Click()
Call SortForm("SortAssignee1", "Assignee")
End Sub

Private Sub SortAssignmentDate1_Click()
Call SortForm("SortAssignmentDate1", "AssignmentDate")
End Sub

Private Sub SortVIN1_Click()
Call SortForm("SortVIN1", "VIN")
End Sub
Private Sub TitleLog1_Click()
Call SortForm("TitleLog1", "TitleLog")
End Sub

Private Sub SortFundDate1_Click()
Call SortForm("SortFundDate1", "FundedDate")
End Sub

Private Sub SortLastEvent1_Click()
Call SortForm("SortLastEvent1", "LastEvent")
End Sub

Private Sub SortUnitDesc1_Click()
Call SortForm("SortUnitDesc1", "UnitDesc")
End Sub

Private Sub SortUnitStatus1_Click()
Call SortForm("SortUnitStatus1", "UnitStatus")
End Sub
Private Sub SortLatestDelvDate1_Click()
Call SortForm("SortLatestDelvDate1", "LatestDelvDate")
End Sub
Private Sub SortDaysElapsed1_Click()
Call SortForm("SortDaysElapsed1", "DaysElapsed")
End Sub
Private Sub SortAssetID1_Click()
Call SortForm("SortAssetID1", "AssetID")
End Sub

Private Sub SortClient1_Click()
Call SortForm("SortClient1", "Client")
End Sub
Private Sub SortUnitGroup1_Click()
Call SortForm("SortUnitGroup1", "UnitGroup")
End Sub
Private Sub SortSchNo1_Click()
Call SortForm("SortSchNo1", "Schedule")
End Sub
Public Function SortForm(CtrlNameX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Sort*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me.Form.RecordSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me.Form.RecordSource = SQLz
End Function
Public Sub ClearSorts(WhichSorts As Integer)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Sort*" Then
            If WhichSorts = 1 And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If WhichSorts = 2 And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
            If WhichSorts = 0 Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
End Sub
