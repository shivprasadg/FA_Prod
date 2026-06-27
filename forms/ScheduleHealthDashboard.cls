Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim RST        As DAO.Recordset
Dim SQLi, SQLx, SQLy, SQLz, TargetQuery As String
Dim CGrpInt As Long

Private Sub ClearFilters_Click()
    Me.SchSelection = 3
    Call SchSelection_Click
    ''Call SearchLessee_Click
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub
Private Sub Excel11_Click()
Call SchHealthReport("11X17", TargetQuery)
End Sub

Private Sub ExcelLegal_Click()
Call SchHealthReport("Legal", TargetQuery)
End Sub


Private Sub FATitleTrust_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub

Private Sub FinalEconRecv_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub

Private Sub Form_Current()
    Me.txtCurrent = Me.SchID
End Sub

Public Sub Form_Load()
    Me.SearchSchNo = ""
    Me.SearchSchOutBack = ""
    Me.SearchRALOutBack = ""
    Me.SearchLessee = ""
ClearSorts (1)
Me.Form.RecordSource = "Select * FROM " & TargetQuery & " "
Me.SearchLessee.RowSource = "Select Lessee From " & TargetQuery & " Group by Lessee Order By Lessee"
Me.SearchSchNo.RowSource = "SELECT Schno FROM " & TargetQuery & " GROUP BY Schno, Lessee ORDER BY Schno;"
End Sub

Private Sub Form_Open(Cancel As Integer)
TargetQuery = "ScheduleHealthCheckUnFunded"
End Sub

Private Sub Lessee_DblClick(Cancel As Integer)
Call OpenClientsForm
End Sub
Private Sub FundDate_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub AssignmentDate_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub Assignee_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub BLCD_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub SchNo_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub
Private Sub RALOutBack_DblClick(Cancel As Integer)
Call OpenGroupForm
End Sub
Private Sub SchOutBack_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub
Private Sub PermPkgOut_DblClick(Cancel As Integer)
Call OpenScheduleForm
End Sub
Private Sub OPSPkg_DblClick(Cancel As Integer)
Call OpenGroupForm
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
Private Sub EarliestDelvDate_DblClick(Cancel As Integer)
If IsDate(EarliestDelvDate) Then
Call OpenMassMod(DLookup("SgrpID", "vw_SixKeys", "SchID=" & Me.SchID & " AND ActualDeliveryDate= #" & Me.EarliestDelvDate & "#"))
Else
Call OpenMassMod(DLookup("SgrpID", "vw_SixKeys", "SchID=" & Me.SchID))
End If
Form_BulkChange_Form.Rents.SetFocus
End Sub
Private Sub LatestDelvDate_DblClick(Cancel As Integer)
If IsDate(LatestDelvDate) Then
Call OpenMassMod(DLookup("SgrpID", "vw_SixKeys", "SchID=" & Me.SchID & " AND ActualDeliveryDate= #" & Me.LatestDelvDate & "#"))
Else
Call OpenMassMod(DLookup("SgrpID", "vw_SixKeys", "SchID=" & Me.SchID))
End If
Form_BulkChange_Form.Rents.SetFocus
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

Private Sub SchSelection_Click()
    If Me.SchSelection = 1 Then TargetQuery = "ScheduleHealthCheck"
    If Me.SchSelection = 2 Then TargetQuery = "ScheduleHealthCheckFunded"
    If Me.SchSelection = 3 Then TargetQuery = "ScheduleHealthCheckUnFunded"
    Call Form_Load
End Sub

Private Sub SearchLessee_Click()
    Me.SearchSchNo = ""
    Me.SearchSchOutBack = ""
    Me.SearchRALOutBack = ""
    ClearSorts (1)
    SQLz = "Select * FROM " & TargetQuery & " "
    SQLz = SQLz & " WHERE Lessee Like '*" & Me.SearchLessee.Column(0) & "*' "
    SQLz = SQLz & "Order By Lessee"
    Me.Form.RecordSource = SQLz
    Me.SearchSchNo.RowSource = "SELECT Schno FROM " & TargetQuery & " WHERE Lessee  Like '*" & Me.SearchLessee.Column(0) & "*' GROUP BY Schno, Lessee ORDER BY Schno;"
End Sub

Private Sub SearchSchNo_Click()
    ''Me.SearchLessee = ""
    Me.SearchSchOutBack = ""
    Me.SearchRALOutBack = ""
   ' ClearSorts (1)
    SQLz = "Select * FROM " & TargetQuery & " "
    SQLz = SQLz & " WHERE Lessee  Like '*" & Me.SearchLessee.Column(0) & "*'  AND SchNo Like '*" & Me.SearchSchNo.Column(0) & "*' "
    SQLz = SQLz & "Order By SchNo"
    Me.Form.RecordSource = SQLz
End Sub
Private Sub SearchSchOutBack_Click()
    Me.SearchLessee = ""
    Me.SearchSchNo = ""
    Me.SearchRALOutBack = ""
    ClearSorts (1)
    SQLz = "Select * FROM " & TargetQuery & " "
    SQLz = SQLz & " WHERE SchOutBack Like '*" & Me.SearchSchOutBack.Column(0) & "*' "
    SQLz = SQLz & "Order By SchOutBack"
    Me.Form.RecordSource = SQLz
End Sub
Private Sub SearchRALOutBack_Click()
    Me.SearchLessee = ""
    Me.SearchSchNo = ""
    Me.SearchSchOutBack = ""
    ClearSorts (1)
    SQLz = "Select * FROM " & TargetQuery & " "
    SQLz = SQLz & " WHERE RALOutBack Like '*" & Me.SearchRALOutBack.Column(0) & "*' "
    SQLz = SQLz & "Order By RALOutBack"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub SortAssignee1_Click()
Call SortForm("SortAssignee1", "Assignee")
End Sub

Private Sub SortAssignmentDate1_Click()
Call SortForm("SortAssignmentDate1", "AssignmentDate")
End Sub

Private Sub SortBLCD1_Click()
Call SortForm("SortBLCD1", "BLCD")
End Sub

Private Sub SortEarliestDelvDate1_Click()
Call SortForm("SortEarliestDelvDate1", "EarliestDelvDate")
End Sub

Private Sub SortFundDate1_Click()
Call SortForm("SortFundDate1", "FundDate")
End Sub

Private Sub SortIntermFund1_Click()
Call SortForm("SortIntermFund1", "IntermFund")
End Sub

Private Sub SortIntermFundBank1_Click()
Call SortForm("SortIntermFundBank1", "IntermFundBank")
End Sub

Private Sub SortLatestDelvDate1_Click()
Call SortForm("SortLatestDelvDate1", "LatestDelvDate")
End Sub

Private Sub SortLessee1_Click()
Call SortForm("SortLessee1", "Lessee")
End Sub

Private Sub SortPermPkgOut1_Click()
Call SortForm("SortPermPkgOut1", "PermPkgOut")
End Sub

Private Sub SortRALOutBack1_Click()
Call SortForm("SortRALOutBack1", "RALOutBack")
End Sub
Private Sub SortSchOutBack1_Click()
Call SortForm("SortSchOutBack1", "SchOutBack")
End Sub
Private Sub SortSchNo1_Click()
Call SortForm("SortSchNo1", "SchNo")
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
