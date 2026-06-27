Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Public Ni      As Long
Public ClientIDx, MLIDx, UnitIDx As Long
Public SchIDx As Long
Public GroupIDx As Long
Public KeyX    As String
Public UnitDescX As String
Public ThisTreeForm As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.formName
End Sub

Private Sub Form_Load()
'Set SixTreeView.ImageList = ImageList1
    Me.SearchIT = ""
    Me.SixTreeView.SetFocus
    ThisTreeForm = Me.Form.Name
    Call SetupTreeview
    Call LoadSixTree("FA")
    Call ClearLabel_Click
    Call ClearLabel2_Click
    Me.SelectedNode = ""
    Me.SearchIT = ""
    Me.ScheduleClientInfo_SubForm.Visible = False
    Me.ScheduleInfo_SubForm.Visible = False
    Me.GroupInfo_SubForm.Visible = False
    Me.ScheduleUnitInfo_SubForm.Visible = False
    Me.Units_List.RowSource = ""
    Me.UnitLookUp.RowSource = ""
    Me.AuditListLabel = ""
    Me.AuditList.RowSource = ""
End Sub

Private Sub SetupTreeview()
    With Me.SixTreeView
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .Indentation = 300
        .Appearance = ccFlat
        .HideSelection = False
        .borderStyle = ccFixedSingle
        .HotTracking = True
        .FullRowSelect = True
        .CheckBoxes = False
        .SingleSel = False
        .Sorted = False
        .Scroll = True
        .LabelEdit = tvwManual
        .Font.Name = "Segoe UI"
        .Font.Weight = 500
        .Font.Size = 8.5
    End With
End Sub
Public Function xgetTV() As MSComctlLib.TreeView
  '  Set getTV = Forms(Me.Form.Name).SixTreeView.Object
End Function
Private Sub FindIT_Click()
    If Me.SearchIT <> "" Then Call TreeVeiwFindIT(Me.SearchIT)
    Me.SearchIT.SetFocus
End Sub

Private Sub ImportFAID_Click()
    DoCmd.OpenForm "ATF_SFImporter"
End Sub

Private Sub SearchIT_AfterUpdate()
    If Me.SearchIT <> "" Then Call TreeVeiwFindIT(Me.SearchIT)
    Me.SearchIT.SetFocus
End Sub

Private Sub UnitLookUp_Click()
    If Me.UnitLookUp.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Dim TempKeyX As String
    getTV (Me.Form.Name)
    KeyX = ""
    If Me.UnitSelection <> 1 Then Call LoadSixTree
    TempKeyX = TempKeyX & Me.UnitLookUp.Column(5) & "|"
    TempKeyX = TempKeyX & Me.UnitLookUp.Column(8) & "|"
    TempKeyX = TempKeyX & Me.UnitLookUp.Column(6) & "|"
    TempKeyX = TempKeyX & Me.UnitLookUp.Column(7) & "|"
    TempKeyX = TempKeyX & Me.UnitLookUp.Column(0) & "|"
    KeyX = TempKeyX
    Me.SixTreeView.SetFocus
    ThisTreeForm = Me.Form.Name
    Call TreeVeiwFindKey(TempKeyX)
    Call NodeClickTest
    Call AfterTreeClick
End Sub

Private Sub UnitLookUp_DblClick(Cancel As Integer)
    If Me.UnitLookUp.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitLookUp.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitLookUp.Column(0) & ""
    If Me.UnitSelection > 2 Then Forms!Unitsform.UnitsTabControl.Value = 6
    Me.SearchVin.SetFocus
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
    If Me.Units_List.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.Units_List.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.Units_List.Column(0) & ""
    If Me.UnitSelection > 2 Then Forms!Unitsform.UnitsTabControl.Value = 6
    Me.SearchVin.SetFocus
End Sub

Private Sub ResetForm_Click()
    Me.UnitSelection = 1
    Call Form_Load
End Sub
Public Sub UnitSelection_Click()
    If Me.UnitSelection = 1 Then Call SetupTreeview
    If Me.UnitSelection = 1 Then Call LoadSixTree
    If Me.UnitSelection = 2 Then Call LoadSixTree("FA")
    If Me.UnitSelection = 3 Then Call LoadSixTree("BUnits")
    If Me.UnitSelection = 4 Then Call LoadSixTree("Track")
    If Me.UnitSelection = 5 Then Call LoadSixTree("TRUnits")
    Call ClearLabel_Click
    Call ClearLabel2_Click
    Me.SelectedNode = ""
    Me.SearchIT = ""
End Sub

Public Sub LoadSixTree(Optional FilterX As String)
    On Error GoTo Err_Handler
    Dim SQLz   As String
    Dim dbs    As DAO.DataBase
    Dim RST    As DAO.Recordset
    Dim tv     As MSComctlLib.TreeView
    Dim NodX   As MSComctlLib.Node
    Dim iNode  As Integer
    Dim ClientName, ClientKey As String
    Dim MstName, MstKey, MstPrefix As String
    Dim SchName, SchKey As String
    Dim GrpName, GrpKey As String
    Dim UnitName, UnitKey As String
    Dim ImageX, ImageOpenX As String
    Set tv = getTV(Me.Form.formName)
    tv.Nodes.Clear
    Set dbs = CurrentDb
    SQLz = "SELECT *, ClientShortName as Client, '(' & Status & ') ' & UnitDescShortVer AS unitdesc FROM vw_SixKeys "
    Me.FleetCaption.caption = "All Units"
    Me.UnitSelection = 1
    If FilterX = "FA" Then
        SQLz = SQLz & "WHERE MLOrig ='FA' "
        Me.FleetCaption.caption = "FA Lease Only"
        Me.UnitSelection = 2
    End If
    If FilterX = "BUnits" Then
        SQLz = SQLz & "WHERE Status ='B' "
        Me.FleetCaption.caption = "B Units Only"
        Me.UnitSelection = 3
    End If
    If FilterX = "Track" Then
        SQLz = SQLz & "WHERE MLOrig <> 'FA' "
        Me.FleetCaption.caption = "Tracking Only"
        Me.UnitSelection = 4
    End If
    If FilterX = "TRUnits" Then
        SQLz = SQLz & "WHERE Status ='TR' "
        Me.FleetCaption.caption = "TR Units Only"
        Me.UnitSelection = 5
    End If
    SQLz = SQLz & "ORDER BY ClientShortName, MLorig, MLNO, SchID Desc, SGrpID;"
    Set RST = dbs.OpenRecordset(SQLz, dbOpenSnapshot)
    Set NodX = tv.Nodes.Add(, , ClientKey, ClientName)

    If Not RST.BOF Then RST.MoveFirst
    Do Until RST.EOF
        ' Add Client nodes
        If ClientName <> RST!Client Then
            ClientKey = RST!clientID & "|"
            ClientName = ProperCase(RST!Client, 0)
            Set NodX = tv.Nodes.Add(, , ClientKey, ClientName, "Client", "ClientOpen")
            NodX.Tag = "Client:" & RST!clientID
            MstName = ""
        End If
        ' Add Lease node
        If MstName <> RST!MLNo Then
            MstKey = ClientKey & RST!MLID & "|"
            MstName = RST!MLNo
            If RST!MLOrig = "FA" Then
                MstPrefix = "Lease "
                ImageX = "Lease"
                ImageOpenX = "LeaseOpen"
            Else
                MstPrefix = "Trk "
                ImageX = "Tracking"
                ImageOpenX = "TrackingOpen"
            End If
            Set NodX = tv.Nodes.Add(ClientKey, tvwChild, MstKey, MstName, ImageX, ImageOpenX)
            NodX.Tag = ImageX & ":" & RST!MLID
            SchName = ""
        End If
        ' Add Schedule node
        On Error Resume Next
        If SchName <> RST!Schedule Then
            SchKey = MstKey & RST!SchID & "|"
            SchName = RST!Schedule
            ImageX = "Schedule"
            Set NodX = tv.Nodes.Add(MstKey, tvwChild, SchKey, "Schedule " & SchName, ImageX, "FolderOpen")
            NodX.Tag = "Schedule:" & RST!SchID
            'nodX.Bold = True
            GrpName = ""
        End If
        On Error GoTo 0
        'Add Group node
        On Error Resume Next
        If GrpName <> RST!UnitGroup & ", " & RST!AssetsInGroup & " units" Then
            GrpKey = SchKey & RST!SGrpID & "|"
            GrpName = RST!UnitGroup & ", " & RST!AssetsInGroup & " units"
            ImageX = "GroupName"
            Set NodX = tv.Nodes.Add(SchKey, tvwChild, GrpKey, "Group " & GrpName, ImageX, "GroupNameOpen")
            NodX.Tag = "Group:" & RST!SGrpID
            NodX.ForeColor = RGB(0, 102, 153)
            ' nodX.Bold = True
            UnitName = ""
        End If
        On Error GoTo 0
        'Add Unit node
        On Error Resume Next
        '  If UnitName <> RST!AssetID & " " & RST!unitdesc Then
        '      UnitKey = GrpKey & RST!AssetID & "|"
        '      UnitName = RST!AssetID & " " & RST!unitdesc
        If UnitName <> RST!unitdesc Then
            UnitKey = GrpKey & RST!unitdesc & "|"
            UnitName = RST!unitdesc
            ImageX = "Truck"
            If RST!unitdesc Like "*Daycab*" Then ImageX = "DayCab"
            If RST!unitdesc Like "*Sleeper*" Then ImageX = "Sleeper"
            If RST!unitdesc Like "*Trailer*" Then ImageX = "Trailer"
            If RST!unitdesc Like "*Auto*" Then ImageX = "Trailer"
            If RST!unitdesc Like "*Straight*" Then ImageX = "Straight"
            If RST!unitdesc Like "*Warrant*" Then ImageX = "Doc"
            Set NodX = tv.Nodes.Add(GrpKey, tvwChild, UnitKey, UnitName, ImageX, ImageX)
            '     NodX.Tag = "Unit:" & RST!AssetID
            ' If iNode = 0 Then iNode = 1 Else iNode = 0
            'If iNode = 1 Then nodX.BackColor = RGB(255, 255, 220) Else nodX.BackColor = RGB(255, 255, 255)
        End If
        On Error GoTo 0
MoveOn:
        RST.MoveNext
    Loop
    Me.Units_List.RowSource = ""
    Me.UnitLookUp.RowSource = ""
    Call ClearLabel_Click
    Call ClearLabel2_Click
    Me.SelectedNode = ""
    Me.SearchIT = ""
Exit_Here:
    Set dbs = Nothing
    Set RST = Nothing
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion
    'Resume Next
End Sub

Private Sub LoadMyListView(ByVal au_id As String)
    On Error GoTo Err_Handler
    Dim SQLz   As String
    Dim dbs    As DAO.DataBase
    Dim RST    As DAO.Recordset
    Dim lstItem As ListItem
    Dim strItem As String
    Dim strKey As String
    Dim intCount As Integer


    ' Remove all existing items
    While Me!lvwMyListView.ListItems.Count > 0
        Me!lvwMyListView.ListItems.Remove (1)
    Wend

    'open recordset of filtered Title/Authors
    Set dbs = CurrentDb
    SQLz = "SELECT title, title_id FROM qryTitleAuthor WHERE [au_id]='" & au_id & "'"
    Set RST = dbs.OpenRecordset(SQLz, dbOpenSnapshot)

    'add each entry in the rs to the listview
    If Not RST.BOF Then RST.MoveFirst
    Do Until RST.EOF
        strItem = RST!Title
        strKey = "ID=" & RST!title_id

        Set lstItem = Me!lvwMyListView.ListItems.Add(, strKey, strItem)

        intCount = intCount + 1
        RST.MoveNext
    Loop

    If intCount > 0 Then Me!lvwMyListView.ListItems(1).Selected = True

Exit_Here:
    Set dbs = Nothing
    Set RST = Nothing
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "ERROR"
    Resume Next
End Sub

Private Sub ReportsList_Click()
    Dim FormX, CommandX As String
    FormX = Me.ReportsList.Column(3)
    CommandX = Me.ReportsList.Column(2)
    If Me.SchID & 99 = 99 Then CommandX = "XXXX"
    Select Case CommandX
    Case "ScheduleCheckLists_Click"
        Call MakeScheduleWorkTable(Me.SchID)
        DoCmd.OpenReport "ScheduleCheckListNew", acViewPreview, , "SchID=" & Me.SchID, acNormal
    Case "SerialNumbers_Click"
        Call MakeScheduleWorkTable(Me.SchID)
        DoCmd.OpenReport "ScheduleSerialNumbers", acViewPreview, "", "SchID=" & Me.SchID, acNormal
    Case "SchCheck_Click"
        If Me.clientID = 4029 Then ReportX = "SchedCheckTLS" Else ReportX = "SchedCheck2"
        DoCmd.close acForm, ReportX
        DoCmd.OpenReport ReportX, acViewPreview, "", "SchID=" & Me.SchID
    Case "DocsUnitList_Click"
        Call MakeScheduleWorkTable(Me.SchID)
        Call MakeScheduleReport(Me.SchID)
    Case "OpsUnitList_Click"
        Call MakeScheduleWorkTable(Me.SchID)
        Call MakeOPSReport(, Me.SchID)
    Case "Custodian_Click"
        Call MakeScheduleWorkTable(Me.SchID)
        Call MakeCustodianSubmission(Me.SchID)
    Case "ClientUnitSummary_Click"
        Call MakeClientUnitSummary(Me.clientID)
    Case Else
        MsgBox "No Schedule or Group Selected", vbInformation, AtlaasVersion
    End Select

End Sub
Private Sub SearchAssetId_Change()
' Me.SearchAssetId = ""
    Me.SearchVin = ""
    Me.SearchUnit = ""
    Me.SearchMYear = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, Make, MYear, SchID FROM vw_SixKeys " & _
           "Where SchID=" & Me.SchID & " AND AssetID Like '*" & Me.SearchAssetId.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
End Sub
Private Sub SearchVIN_Change()
    Me.SearchAssetId = ""
    ' Me.SearchVin = ""
    Me.SearchUnit = ""
    Me.SearchMYear = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, Make, MYear, SchID FROM vw_SixKeys " & _
           "Where SchID=" & Me.SchID & " AND VIN Like '*" & Me.SearchVin.Text & "*' Order BY VIN, AssetId;"
    Me.Units_List.RowSource = SQLz
End Sub

Private Sub SearchUnit_Change()
    Me.SearchAssetId = ""
    Me.SearchVin = ""
    ' Me.SearchUnit = ""
    Me.SearchMYear = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, Make, MYear, SchID FROM vw_SixKeys " & _
           "Where SchID=" & Me.SchID & " AND UnitNum Like '*" & Me.SearchUnit.Text & "*' Order BY UnitNum, AssetId;"
    Me.Units_List.RowSource = SQLz
End Sub

Private Sub SearchMYear_Change()
    Me.SearchAssetId = ""
    Me.SearchVin = ""
    Me.SearchUnit = ""
    'Me.SearchMYear = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, Make, MYear, SchID FROM vw_SixKeys " & _
           "Where SchID=" & Me.SchID & " AND MYear Like '*" & Me.SearchMYear.Text & "*' Order BY MYear, Make, AssetId;"
    Me.Units_List.RowSource = SQLz
End Sub
Private Sub AssetSearch2_Change()
'Me.AssetSearch2 = ""
    Me.VINSearch2 = ""
    Me.UnitNumSearch2 = ""
    Me.PlateNumSearch2 = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, UnitPlateNum, ClientShortName, clientID, SchID, SGrpID, MLID FROM vw_SixKeys " & _
           "Where AssetId Like '*" & Me.AssetSearch2.Text & "*' Order By AssetID;"
    Me.UnitLookUp.RowSource = SQLz
End Sub

Private Sub ViewATF_Click()
    DoCmd.close acForm, "ATFTreeCommand"
    DoCmd.OpenForm "ATFTreeCommand"

End Sub

Private Sub VinSearch2_Change()
    Me.AssetSearch2 = ""
    'Me.VINSearch2 = ""
    Me.UnitNumSearch2 = ""
    Me.PlateNumSearch2 = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, UnitPlateNum, ClientShortName, clientID, SchID, SGrpID, MLID FROM vw_SixKeys " & _
           "Where AssetId Like '*" & Me.VINSearch2.Text & "*' Order By Vin, AssetID;"
    Me.UnitLookUp.RowSource = SQLz
End Sub
Private Sub UnitNumSearch2_Change()
    Me.AssetSearch2 = ""
    Me.VINSearch2 = ""
    'Me.UnitNumSearch2 = ""
    Me.PlateNumSearch2 = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, UnitPlateNum, ClientShortName, clientID, SchID, SGrpID, MLID FROM vw_SixKeys " & _
           "Where AssetId Like '*" & Me.UnitNumSearch2.Text & "*' Order By UnitNum, AssetID;"
    Me.UnitLookUp.RowSource = SQLz
End Sub
Private Sub PlateNumSearch2_Change()
    Me.AssetSearch2 = ""
    Me.VINSearch2 = ""
    Me.UnitNumSearch2 = ""
    'Me.PlateNumSearch2 = ""
    SQLz = "SELECT AssetID, VIN, UnitNum, UnitPlateNum, ClientShortName, clientID, SchID, SGrpID, MLID FROM vw_SixKeys " & _
           "Where UnitPlateNum Like '*" & Me.PlateNumSearch2.Text & "*' Order By AssetID;"
    Me.UnitLookUp.RowSource = SQLz
End Sub
Private Sub ClearLabel_Click()
    Me.SearchVin = ""
    Me.SearchUnit = ""
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub
Private Sub ClearLabel2_Click()
    Me.AssetSearch2 = ""
    Me.PlateNumSearch2 = ""
    Me.UnitNumSearch2 = ""
    Me.VINSearch2 = ""
    Me.VINSearch2.SetFocus
    Me.UnitLookUp.RowSource = ""
End Sub

Private Sub SixTreeView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Long, ByVal Y As Long)
    Dim NodX   As Node
    Set NodX = getTV(Me.Form.formName).HitTest(x, Y)
    ' On Error GoTo NoNodeClicked
    If Button = eMouse.Rightclick Then
        If NodX Is Nothing Then
            Call NodeClickTest
            Call AfterTreeClick
            Call RightClickEmptySpace(Me.Form.Name)
        Else
            Call RightClickNode(Me.Form.Name, NodX.Tag)
            'MsgBox nodX.Text & vbNewLine & nodX.Tag
        End If
    End If
    If Button = eMouse.LeftClick Then
        Call NodeClickTest
        Call AfterTreeClick
    End If
    '  If nodX.Tag Like "Client*" Then Call ShowClientInfo(CLng(Split(nodX.Tag, ":")(1)))

    '  If Button = acRightButton Then
    Exit Sub
NoNodeClicked:
    Me.ScheduleClientInfo_SubForm.Form.RecordSource = ""
    Me.ScheduleClientInfo_SubForm.Form.Visible = False
End Sub
Public Sub ClearBoxes()
        Me.AuditListLabel = ""
        Me.AuditList.RowSource = ""
        Me.WhoCbo = ""
        Me.WhatCbo = ""
        Me.WhoCbo.RowSource = ""
        Me.WhatCbo.RowSource = ""
End Sub

Private Sub AfterTreeClick()
'Click on Tree Event
Call ClearBoxes
Call SetCbo(1)
    If ClientIDx <> 0 Then
        If Nz(Me.clientID, 0) <> ClientIDx Then Call ShowClientInfo
        Me.clientID = ClientIDx
        Me.ClientGroupID = DLookup("ClientGroupID", "vw_SixKeys", "ClientId=" & Me.clientID)
        Me.SelectedNode = "Co:<B>" & DLookup("ClientShortName", "vw_SixKeys", "ClientId=" & Me.clientID)
    Else
        Me.ScheduleClientInfo_SubForm.Form.Visible = False
        Me.clientID = 0
        Me.ClientGroupID = 0
    End If
    If MLIDx <> 0 Then
        Me.MLID = MLIDx
        Me.SelectedNode = Me.SelectedNode & " </B>ML:<B> " & DLookup("MLNO", "vw_SixKeys", "MLID=" & Me.MLID)
        Me.ScheduleClientInfo_SubForm.Form.MLID = Me.MLID
        Me.ScheduleClientInfo_SubForm.Form.MasterLease = DLookup("MLNO", "vw_SixKeys", "MLID=" & Me.MLID)
    End If
    If SchIDx <> 0 Then
        If Nz(Me.SchID, 0) <> SchIDx Then Call ShowScheduleInfo
        Me.SchID = SchIDx
        Me.SelectedNode = Me.SelectedNode & " </B>Sch:<B> " & DLookup("Schedule", "vw_SixKeys", "SchID=" & Me.SchID)
    Else
        Me.ScheduleInfo_SubForm.Form.Visible = False
        Me.SchID = 0
    End If
    If GroupIDx <> 0 Then
        If Nz(Me.groupID, 0) <> GroupIDx Then Call ShowGroupInfo
        Me.groupID = GroupIDx
        Me.SelectedNode = Me.SelectedNode & " </B>Grp:<B> " & DLookup("unitGroup", "vw_SixKeys", "SgrpID=" & Me.groupID)
    Else
        '   Me.GroupInfo_SubForm.Form.Visible = False
        Me.groupID = 0
    End If
    If UnitIDx <> 0 Then
        '       Me.Units_List = UnitIDx
        '        Me.SelectedNode = Me.SelectedNode & " </B>AssetID:<B> " & DLookup("AssetId", "vw_SixKeys", "AssetID=" & Me.UnitIDx)
        '       Call Units_List_Click
    End If
    Call SortListBox("ToggleSortWhat1", "AuditList", "FieldLabel", 1)
End Sub
Private Sub ShowClientInfo()
    Me.ScheduleClientInfo_SubForm.Form.Visible = True
    SQLz = "Select * From ScheduleClientInfo Where ClientId=" & ClientIDx & ";"
    Me.ScheduleClientInfo_SubForm.Form.RecordSource = SQLz
End Sub
Private Sub ShowScheduleInfo()
    Me.ScheduleInfo_SubForm.Form.Visible = True
    SQLz = "Select * From ScheduleInfo Where SchID =" & SchIDx & ";"
    Me.ScheduleInfo_SubForm.Form.RecordSource = SQLz
End Sub
Private Sub ShowGroupInfo()
    Me.AuditListLabel = "<b>Group</b> Changes"
    Me.GroupInfo_SubForm.Form.Visible = True
    SQLz = "SELECT *, DLookUp('MLOrig','vw_SixKeys','SGrpID =' & SGrpID) As MLOrig FROM Schgrp WHERE Schgrp.SGrpID=" & GroupIDx & ";"
    Me.GroupInfo_SubForm.Form.RecordSource = SQLz
    Me.groupID = GroupIDx
    SQLz = "SELECT AssetID, VIN, UnitNum,  Make, MYear, SGrpID " & _
           "From vw_SixKeys Where SGrpID=" & GroupIDx & " Order By AssetID;"
    Me.Units_List.RowSource = SQLz
End Sub
Public Sub NodeClickTest()
    Dim TmpStr As String
    On Error Resume Next
    KeyX = Me.SixTreeView.SelectedItem.Key
    Ni = CountStringOccurances(KeyX, "|")
    On Error GoTo 0
    UnitIDx = 0
    GroupIDx = 0
    SchIDx = 0
    MLIDx = 0
    ClientIDx = 0
    If Ni < 4 Then Me.Units_List.RowSource = ""
    If Ni = 5 Then
        Ni = Ni - 1
        TmpStr = (Split(KeyX, "|")(Ni))
        'If TmpStr <> "" Then UnitIDx = CLng(Split(TmpStr, "|")(0))
        If TmpStr <> "" Then UnitDescX = SplitString(TmpStr, "|", 0)
    End If
    If Ni = 4 Then
        Ni = Ni - 1
        TmpStr = (Split(KeyX, "|")(Ni))
        GroupIDx = CLng(Split(TmpStr, "|")(0))
    End If
    If Ni = 3 Then
        Ni = Ni - 1
        TmpStr = (Split(KeyX, "|")(Ni))
        SchIDx = CLng(Split(TmpStr, "|")(0))
    End If
    If Ni = 2 Then
        Ni = Ni - 1
        TmpStr = (Split(KeyX, "|")(Ni))
        MLIDx = CLng(Split(TmpStr, "|")(0))
    End If
    If Ni = 1 Then
        Ni = Ni - 1
        TmpStr = (Split(KeyX, "|")(Ni))
        ClientIDx = CLng(Split(TmpStr, "|")(0))
    End If
End Sub
Private Sub ToggleSortWho1_Click()
    Call SortListBox("ToggleSortWho1", "AuditList", "SysUser")
End Sub
Private Sub ToggleSortWhat1_Click()
    Call SortListBox("ToggleSortWhat1", "AuditList", "FieldLabel")
End Sub
Private Sub ToggleSortFrom1_Click()
    Call SortListBox("ToggleSortFrom1", "AuditList", "oldValue")
End Sub
Private Sub ToggleSortTo1_Click()
    Call SortListBox("ToggleSortTo1", "AuditList", "newValue")
End Sub
Private Sub ToggleSortWhen1_Click()
    Call SortListBox("ToggleSortWhen1", "AuditList", "ActionDate")
End Sub

Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String, Optional Direction As Boolean)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 And Direction = False Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Private Sub Units_List_Click()
    If Me.Units_List.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    SQLz = "SELECT * From Units WHERE UnitID=" & Me.Units_List.Column(0) & ";"
    Me.ScheduleUnitInfo_SubForm.Form.RecordSource = SQLz
    'Me.ScheduleUnitInfo_SubForm!VendorName = DLookup("[VendorName]", "Vendors", "VendorID=" & DLookup("VendorID", "vw_SixKeys", "SGrpId=" & Me.groupID))
    Me.ScheduleUnitInfo_SubForm.Visible = True
    Me.AuditListLabel = "<b>Unit</b> Changes"
    Call SetCbo(1, 1)
End Sub
Public Sub SetCbo(Optional Clearcbo As Boolean, Optional IsUnit As Boolean)
   Dim TableX As String
   Dim IDX As Long
   If Clearcbo = True Then
     Me.WhoCbo = ""
     Me.WhatCbo = ""
   End If
    If CountStringOccurances(KeyX, "|") = 3 Or Me.AuditListLabel Like "*Schedule*" Then
      TableX = "Schedule"
      IDX = Me.SchID
      Me.AuditListLabel = "<b>Schedule</b> Changes"
   End If
      'If Me.AuditListLabel Like "*Group*" Then
      If CountStringOccurances(KeyX, "|") = 4 Or Me.AuditListLabel Like "*Group*" Then
      TableX = "SchGrp"
      IDX = Me.groupID
      Me.AuditListLabel = "<b>Group</b> Changes"
   End If
      If IsUnit = True Or Me.AuditListLabel Like "*Unit*" Then
      TableX = "Units"
      Me.AuditListLabel = "<b>Unit</b> Changes"
   End If
    SQLz = "SELECT identity, SysUser AS Who, FieldLabel AS What, oldValue AS [From], newValue AS [To] ,ActionDate AS [When], tableName AS [Where], Action " & _
           "FROM vw_FAID_RevBase as Audit Where Action <> 'Added' and tablename='" & TableX & "' and Identity=" & IDX & " Order By FieldLabel, ActionDate Desc;"
    Me.AuditList.RowSource = SQLz
    SQLz = "SELECT SysUser AS Who FROM vw_FAID_RevBase as Audit " & _
           "Where Action <> 'Added' and tablename='" & TableX & "' and Identity=" & IDX & " " & _
           "Group By SysUser Order By SysUser;"
    Me.WhoCbo.RowSource = SQLz
    SQLz = "SELECT FieldLabel AS What FROM vw_FAID_RevBase as Audit " & _
           "Where Action <> 'Added' and tablename='" & TableX & "' and Identity=" & IDX & " " & _
           "Group By FieldLabel Order By FieldLabel;"
    Me.WhatCbo.RowSource = SQLz
End Sub

Private Sub WhoCbo_Click()
Me.WhatCbo = ""
Call SetCbo
SQLz = SplitString(Me.AuditList.RowSource, " Order By", 0) & " AND SysUser='" & Me.WhoCbo & "' Order By FieldLabel, ActionDate Desc;"
Me.AuditList.RowSource = SQLz
End Sub
Private Sub WhatCbo_Click()
Me.WhoCbo = ""
Call SetCbo
SQLz = SplitString(Me.AuditList.RowSource, " Order By", 0) & " AND SysUser='" & Me.WhatCbo & "' Order By FieldLabel, ActionDate Desc;"
Me.AuditList.RowSource = SQLz
End Sub
