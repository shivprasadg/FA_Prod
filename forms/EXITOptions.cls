Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, SQLu, ActionType, Response As String
Dim UnitIDz As Long
Dim RST As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

'Private Sub ClearFilters_Click()
'    Me.ExitOptions_subform_Rents.Form.FilterOn = False
'End Sub

Private Sub CloseMe_Click()
    On Error GoTo FixMe
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
FixMe:
End Sub



Public Sub ExchangeSelected_Click()
    Call ConfrimIT("ExchangeIT")
    If Response = vbNo Then Exit Sub
    Call PlaceSelectedUnits("ExchangeIT")
    Me.Tabcontrol.Value = 0
    Call GetUnitCounts
    SQLz = "Select * from ExitOptionsQryExchangeIT Order by UnitID;"
    Me.ExitOptions_subform_Exchange.Form.RecordSource = SQLz
    Me.ExitOptions_subform_Exchange.Form.LeasePaidAdvArr = DLookup("LeasePaidAdvArr", "SchGrp", "sGrpID=" & Me.GroupList.Column(0))
    Me.UnitCount.ControlSource = "=ExitOptions_subform_Exchange.Form!UnitCount"
End Sub

Private Sub ExitOptionsConsole_Click()
    DoCmd.close acForm, "ExitOptionsManager"
    DoCmd.OpenForm "ExitOptionsManager"
    If Me.ClientList.Column(0) & "test" <> "test" Then
        Forms!ExitOptionsManager.ClientList = Me.ClientList.Column(0)
        Form_ExitOptionsManager.ClientList_Click
        If Me.GroupList.Column(0) & "test" <> "test" Then
            Forms!ExitOptionsManager.SchGrpList = Me.GroupList.Column(0)
            Form_ExitOptionsManager.SchGrpList_Click
        End If
    End If
End Sub

Public Sub ExtendSelected_Click()
    Call ConfrimIT("Extend Lease")
    If Response = vbNo Then Exit Sub
    Call PlaceSelectedUnits("Extend")
    ActionType = "Extend"
    Me.Tabcontrol.Value = 1
    SQLz = "Select * From ExitOptionsQryExtend Order By UnitID;"
    Me.ExitOptions_subform_Extend.Form.RecordSource = SQLz
    Call GetUnitCounts
    Me.UnitCount.ControlSource = "=ExitOptions_subform_Extend.Form!UnitCount"
End Sub

Private Sub Form_Current()
    Me.Units_List.ControlTipText = Nz(Units_List.Column(3), "Click Unit to see Unit's Description")
End Sub
Public Sub BuyITSelection_Click()
    Call ConfrimIT("Buy-IT")
    If Response = vbNo Then Exit Sub
    Call PlaceSelectedUnits("BuyIT")
    ActionType = "BuyIT"
    Me.Tabcontrol.Value = 4
    SQLz = "Select * From ExitOptionsQryBuyIT Order by UnitID;"
    Me.ExitOptions_subform_BuyIT.Form.RecordSource = SQLz
    Call GetUnitCounts
    Me.UnitCount.ControlSource = "=ExitOptions_subform_BuyIT.Form!UnitCount"
End Sub

Private Sub GetLatestData_Click()
    Call MakeExitOptionLocal(True)
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT ClientID, Client, [Type], Units, ClientGroupID " & _
           "FROM ClientMasterLeasesTypeCount_Summed Where Client Like '*" & Me.SearchClient.Text & "*' " & _
           "ORDER BY Client;"
    Me.ClientList.RowSource = SQLz
End Sub

Public Sub TerminateSelected_Click()
    Call ConfrimIT("Terminate")
    If Response = vbNo Then Exit Sub
    Call PlaceSelectedUnits("Terminate")
    ActionType = "Terminate"
    Me.Tabcontrol.Value = 2
    Call GetUnitCounts
    SQLz = "Select * From ExitOptionsQryTerminate Order by UnitID;"
    Me.ExitOptions_subform_Terminate.Form.RecordSource = SQLz
    Me.UnitCount.ControlSource = "=ExitOptions_subform_Terminate.Form!UnitCount"
End Sub
Public Sub UnitHistory_click()
    If Me.Units_List & "test" = "test" And IsNull(UnitIDz) Then Exit Sub
    Call PlaceSelectedUnits("History")
    Call GetUnitCounts
    Me.Tabcontrol.Value = 3
    Me.ExitOptions_subform_History.Visible = True
    Me.UnitCount.ControlSource = "=ExitOptions_subform_History.Form!UnitCount"
End Sub

Public Sub History_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    Call GetAll_Click
    Call PlaceSelectedUnits("History")
    Call GetUnitCounts
    Me.Tabcontrol.Value = 3
    SQLz = "SELECT vw_ExitOptionHistory.* FROM vw_ExitOptionHistory " & _
           "INNER JOIN EXITOptions_WorkTable ON vw_ExitOptionHistory.UnitID = EXITOptions_WorkTable.UnitID;"
    Me.ExitOptions_subform_History.Form.RecordSource = SQLz
    Me.ExitOptions_subform_History.Visible = True
    Me.UnitCount.ControlSource = "=ExitOptions_subform_History.Form!UnitCount"
End Sub
Public Sub ConfrimIT(ActionX As String)
    Dim Msgx As String
    Msgx = "You're about to Add/Modify the" & vbNewLine & ActionX
    Msgx = Msgx & " Options for " & ([Units_List].[ItemsSelected].[Count]) & ""
    Msgx = Msgx & " Unit"
    If ([Units_List].[ItemsSelected].[Count]) > 1 Then Msgx = Msgx & "s"
    Msgx = Msgx & vbNewLine & vbNewLine & "Are you sure?"
    If ActionX Like "*Buy*" Then Msgx = "You are about Buy Portfolio Units and link them to a New Lease." & vbNewLine & "Do you want to continue?"
    Response = MsgBox(Msgx, vbQuestion + vbYesNo, "Add/Modify Warning")
End Sub
Public Sub Form_Load()
    OptionType = ""
    Me.SearchtDesc = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Call SetSQLu
    Call DelWorkTable
    Call ClearBoxes
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.UnitCount.Visible = False
End Sub
Public Sub SetSQLu()
    SQLu = "Select Distinct"
    SQLu = SQLu & " AssetID"
    SQLu = SQLu & ",uStatus As S"
    SQLu = SQLu & ",VIN"
    SQLu = SQLu & ",UnitNum"
    SQLu = SQLu & ",Selected"
    SQLu = SQLu & ",UnitDesc"
    SQLu = SQLu & ",Format(ExitOptionsGroup_Local.EOLPrimary,'mm/dd/yy') as EOLPrimary"
    SQLu = SQLu & ",OffLeaseProj "
    SQLu = SQLu & " FROM ExitOptionsGroup_Local "
End Sub
Private Sub GetAll_Click()
    For i = 0 To Units_List.ListCount - 1
        Units_List.Selected(i) = True
    Next i
    Call GetUnitCounts
End Sub

Private Sub GetSelectedUnits_Click()
    Call PlaceSelectedUnits("ExchangeIT")
    Call GetUnitCounts
End Sub

Public Sub PlaceSelectedUnits(Typex As String)

    Dim ItemIndex As Variant
    Dim wUnitID As String

    Call DelWorkTable
    DoCmd.SetWarnings False
    For Each ItemIndex In Me.Units_List.ItemsSelected
        SQLz = "INSERT INTO EXITOptions_WorkTable (UnitID, UnitUnitNum, UID0,UID1,UID2, UID3, GroupID ) " & _
               "SELECT " & Me.Units_List.ItemData(ItemIndex) & " AS UnitID" & _
               ", '" & DLookup("UnitUnitNum", "Units", "UnitID=" & Me.Units_List.ItemData(ItemIndex) & "") & "' As UnitNum " & _
               ", '" & Me.Units_List.ItemData(ItemIndex) & "-0' As UID0 " & _
               ", '" & Me.Units_List.ItemData(ItemIndex) & "-1' As UID1 " & _
               ", '" & Me.Units_List.ItemData(ItemIndex) & "-2' As UID2 " & _
               ", '" & Me.Units_List.ItemData(ItemIndex) & "-" & Nz(DMax("OptionTypeID", "UnitExitOption", "OptionTypeID>2 AND UnitID=" & Me.Units_List.ItemData(ItemIndex) & ""), 3) & "' As UID3 " & _
               ",  " & DLookup("GroupRef", "Units", "UnitID=" & Me.Units_List.ItemData(ItemIndex) & "") & " As GroupID;"
        DoCmd.RunSQL (SQLz)
        wUnitID = wUnitID & Me.Units_List.ItemData(ItemIndex) & ","
    Next
    wUnitID = Left(wUnitID, Len(wUnitID) - 1)
    DoCmd.SetWarnings True
  '  Call ChangeConnectionString(cConnection(0), , "UnitExitOption")
    If Typex = "History" Then GoTo ShowForms

    If Typex = "ExchangeIT" Then
        SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX ) " & _
               "SELECT Units.UnitID, 1 AS OptionID, 'ExchangeIT' AS TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, Units.UnitID & '-1' AS UIDX " & _
               "FROM ((Units INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN (Select UnitID From UnitExitOption Where OptionTypeID = 1 AND UnitID in (" & wUnitID & ")) As UEO ON Units.UnitID = UEO.UnitID) " & _
               "WHERE UEO.UnitID Is Null AND Units.UnitID In(" & wUnitID & ");"
       'Changed SQL to Fix Recoord Locking issues on RDP... 3/12/2022 LBJ
       'SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX ) " & _
               "SELECT EXITOptions_WorkTable.UnitID, 1 AS OptionID, 'ExchangeIT' AS TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, EXITOptions_WorkTable.UnitID & '-1' AS UIDX " & _
               "FROM ((EXITOptions_WorkTable INNER JOIN Units ON EXITOptions_WorkTable.UnitID = Units.UnitID) INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN UnitExitOption ON (EXITOptions_WorkTable.UID1 = UnitExitOption.UIDX) AND (EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID) " & _
               "WHERE (((UnitExitOption.UnitID) Is Null));"
        Set RST = CurrentDb.OpenRecordset("ExitOptionCheckExch")

    End If
    If Typex = "Terminate" Then
        SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX,OptionSubType, StartOrPurposedDate ) " & _
               "SELECT Units.UnitID, 2 AS OptionID, 'Terminate' AS TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, Units.UnitID & '-2' AS UIDX, " & _
               "'Normal' As OptionSubType, EOLDate AS EOLDate " & _
               "FROM ((SchGrp INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF) " & _
               "LEFT JOIN (Select UnitID From UnitExitOption Where OptionTypeID = 2 AND UnitID in (" & wUnitID & ")) As UEO ON Units.UnitID = UEO.UnitID) " & _
               "INNER JOIN (SELECT AssetID, EOLDate, ExchDate, EOLDatePrimary, DaysToEOL, DaysToExch, DaysToEOLPrimary " & _
               "FROM LeasedScheduleGroupUnit Where AssetID in (" & wUnitID & ")) As EOL ON Units.UnitID = EOL.AssetID " & _
               "WHERE UEO.UnitID Is Null AND Units.UnitID In(" & wUnitID & ");"
        'Changed SQL to Fix Recoord Locking issues on RDP... 3/12/2022 LBJ
        'SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX,OptionSubType, StartOrPurposedDate ) " & _
               "SELECT EXITOptions_WorkTable.UnitID, 2 AS OptionID, 'Terminate' AS TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, EXITOptions_WorkTable.UnitID & '-2' AS UIDX, " & _
               "'Normal' As OptionSubType, ExitOptions_EOLDates.EOLDate AS EOLDate " & _
               "FROM (((EXITOptions_WorkTable INNER JOIN Units ON EXITOptions_WorkTable.UnitID = Units.UnitID) " & _
               "INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN UnitExitOption ON (EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID) AND (EXITOptions_WorkTable.UID2 = UnitExitOption.UIDX)) " & _
               "INNER JOIN ExitOptions_EOLDates ON EXITOptions_WorkTable.UnitID = ExitOptions_EOLDates.AssetID " & _
               "WHERE (((UnitExitOption.UnitID) Is Null));"

        Set RST = CurrentDb.OpenRecordset("ExitOptionCheckTerm")
    End If

    If Typex = "Extend" Then
        SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX, Rent, TermMonths, Fees, OptionSubType, StartOrPurposedDate, EndOrActualDate, Notes ) " & _
               "SELECT Units.UnitID, 3 AS OptionID, 'Extend' As TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, Units.UnitID & '-3' AS UIDX, " & _
               "Nz(SchGrp.LeaseExtOption1Rent,0) As ExtRent, nz(SchGrp.LeaseExtOption1Term,0) As TermMonths, SchGrp.SgrpFSLPmt As Fees, 'Extend Term' As OptionSubType, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD]) As ExtStartDate, " & _
               "IIf([LeaseExtOption1Term]>0,DateAdd('m',Nz([SchGrp].[LeaseExtOption1Term],0),DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD]))-1,Null) AS ExtLeaseEnd, IIf([LeaseExtOption1Term]>0,'As per Group ExtendIT Options',null) As Notes " & _
               "FROM ((Units INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN (Select UnitID From UnitExitOption Where OptionTypeID = 3 AND UnitID in (" & wUnitID & ")) As UEO ON Units.UnitID = UEO.UnitID) " & _
               "WHERE UEO.UnitID Is Null AND Units.UnitID In(" & wUnitID & ");"
        'Changed SQL to Fix Recoord Locking issues on RDP... 3/12/2022 LBJ
        'SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX, Rent, TermMonths, Fees, OptionSubType, StartOrPurposedDate, EndOrActualDate, Notes ) " & _
               "SELECT EXITOptions_WorkTable.UnitID, 3 AS OptionID, 'Extend' As TypeDesc, SchGrp.BLCD, Units.UnitSynResid, SchGrp.LeaseTermPrimary, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, Units.UnitID & '-3' AS UIDX, " & _
               "Nz(SchGrp.LeaseExtOption1Rent,0) As ExtRent, nz(SchGrp.LeaseExtOption1Term,0) As TermMonths, SchGrp.SgrpFSLPmt As Fees, 'Extend Term' As OptionSubType, DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD]) As ExtStartDate, " & _
               "IIf([LeaseExtOption1Term]>0,DateAdd('m',Nz([SchGrp].[LeaseExtOption1Term],0),DateAdd('m',[LeaseTermPrimary],[SchGrp].[BLCD]))-1,Null) AS ExtLeaseEnd, IIf([LeaseExtOption1Term]>0,'As per Group ExtendIT Options',null) As Notes " & _
               "FROM ((EXITOptions_WorkTable INNER JOIN Units ON EXITOptions_WorkTable.UnitID = Units.UnitID) INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN UnitExitOption ON (EXITOptions_WorkTable.UID3 = UnitExitOption.UIDX) AND (EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID) " & _
               "WHERE (((UnitExitOption.UnitID) Is Null));"
        Set RST = CurrentDb.OpenRecordset("ExitOptionCheckExtd")
    End If

    If Typex = "BuyIT" Then
       SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX ) " & _
               "SELECT Units.UnitID, 0 AS OptionID, 'BuyIT' AS TypeDesc, '1/1/1900', Units.UnitSynResid, 0, '1/1/1900' AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, Units.UnitID & '-0' AS UIDX " & _
               "FROM ((Units INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN (Select UnitID From UnitExitOption Where OptionTypeID = 0 AND UnitID in (" & wUnitID & ")) As UEO ON Units.UnitID = UEO.UnitID) " & _
               "WHERE UEO.UnitID Is Null AND Units.UnitID In(" & wUnitID & ");"
        'Changed SQL to Fix Recoord Locking issues on RDP... 3/12/2022 LBJ
        'SQLz = "INSERT INTO UnitExitOption ( UnitID, OptionTypeID, OptionTypeDesc, OrginalBLCD, OrigResidual, OrigLeaseTerm, OrginalEOLDate, UserName, DatePosted, UIDX ) " & _
               "SELECT EXITOptions_WorkTable.UnitID, 0 AS OptionID, 'BuyIT' AS TypeDesc, '1/1/1900', Units.UnitSynResid, 0, '1/1/1900' AS EOL, '" & FAUserName & "' AS [User], '" & Now() & "'AS Posted, EXITOptions_WorkTable.UnitID & '-0' AS UIDX " & _
               "FROM ((EXITOptions_WorkTable INNER JOIN Units ON EXITOptions_WorkTable.UnitID = Units.UnitID) INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) " & _
               "LEFT JOIN UnitExitOption ON (EXITOptions_WorkTable.UID0 = UnitExitOption.UIDX) AND (EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID) " & _
               "WHERE (((UnitExitOption.UnitID) Is Null));"
        Set RST = CurrentDb.OpenRecordset("ExitOptionCheckBuyIT")

    End If
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If

   ' Call ChangeConnectionString(cConnection(0), , "UnitExitOption")

    'Set GAPLAP data in the following sub
    If Typex = "ExchangeIT" Then Call SetGapLapData(1)
    If Typex = "Terminate" Then SetGapLapData (2)
    If Typex = "Extend" Then Call SetGapLapData(3)
    'Set RST = Nothing
    RST.close
    If Typex = "Terminate" Then
        SQLz = "SELECT DISTINCTROW EXITOptionsGAPReport.AssetID, 2 AS OptionTypeID, IIf([Early]='Yes',-1,0) AS TerminateEarly, EXITOptionsGAPReport.ProjRV AS Residual, " & _
               "EXITOptionsGAPReport.GainLoss AS CostGainOfSale, LastRent As Rent, " & _
               "EXITOptionsGAPReport.OffLeaseDate AS EndOrActualDate, EXITOptionsGAPReport.EOLDateActive as StartOrPurposedDate " & _
               "INTO ExitTermTemp FROM EXITOptions_WorkTable INNER JOIN EXITOptionsGAPReport ON EXITOptions_WorkTable.UnitID = EXITOptionsGAPReport.AssetID;"
        DoCmd.SetWarnings False
        ' removed >>>> , nz(EXITOptionsGAPReport.OffLeaseDate,EOLDateActive) - 30 As FinalRentDate >>> from the above SQLz - 6/1/2016 LBJ
        DoCmd.RunSQL (SQLz)
        SQLz = "UPDATE ExitTermTemp INNER JOIN UnitExitOption ON (ExitTermTemp.OptionTypeID = UnitExitOption.OptionTypeID) AND (ExitTermTemp.AssetID = UnitExitOption.UnitID) " & _
               "SET UnitExitOption.StartOrPurposedDate = [ExitTermTemp].StartOrPurposedDate, UnitExitOption.EndOrActualDate = [ExitTermTemp].[EndOrActualDate], UnitExitOption.Residual = [ExitTermTemp].[Residual], " & _
               "UnitExitOption.CostGainOfSale = [ExitTermTemp].[CostGainOfSale], UnitExitOption.TerminateEarly = [ExitTermTemp].[TerminateEarly], " & _
               "UnitExitOption.Rent = [ExitTermTemp].[Rent];"
        DoCmd.RunSQL (SQLz)
        ' removed >>>> , UnitExitOption.FinalRentDate=[ExitTermTemp].[FinalRentDate] >>> from the above SQLz - 6/1/2016 LBJ

        ' DoCmd.DeleteObject acTable, "ExitTermTemp"
        DoCmd.SetWarnings True
    End If
    ' On Error GoTo fixme1

ShowForms:
    Me.Units_List.SetFocus
    Me.UnitCount.Visible = True

    If Typex = "ExchangeIT" Then
        Me.Tabcontrol.Value = 0
        Me.ExitOptions_subform_Exchange.Visible = True
        Me.ExitOptions_subform_Exchange.Requery
        Me!ExitOptions_subform_Exchange.Form.ClearBoxes
    End If

    If Typex = "Extend" Then
        Me.Tabcontrol.Value = 1
        Me.ExitOptions_subform_Extend.Visible = True
        Me.ExitOptions_subform_Extend.Requery
        Me!ExitOptions_subform_Extend.Form.ClearBoxes
        On Error Resume Next
        Me!ExitOptions_subform_Extend.Form.StartOrPurposedDateCBO = DateAdd("m", DLookup("LeaseTermPrimary", "SchGrp", "SGrpID=" & Me.GroupList.Column(0)), DLookup("BLCD", "SchGrp", "SGrpID=" & Me.GroupList.Column(0)))
        On Error GoTo 0
    End If
    If Typex = "Terminate" Then
        Me.Tabcontrol.Value = 2
        Me.ExitOptions_subform_Terminate.Visible = True
        Me.ExitOptions_subform_Terminate.Requery
        Me!ExitOptions_subform_Terminate.Form.ClearBoxes
    End If
    If Typex = "BuyIT" Then
        Me.Tabcontrol.Value = 4
        Me.ExitOptions_subform_BuyIT.Visible = True
        Me.ExitOptions_subform_BuyIT.Requery
        Me!ExitOptions_subform_BuyIT.Form.ClearBoxes
    End If

    If Typex = "History" Then
        'Me.Tabcontrol.Value = 3
        'Me.ExitOptions_subform_History.Visible = True
         Me.ExitOptions_subform_History.Requery
    End If
    '  Me.Units_List.Requery

    Exit Sub
fixme1:

    On Error GoTo 0
    Exit Sub

GetSelectedUnits_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSelectedUnits_Click of VBA Document Form_ExitOptions"
End Sub
Public Sub SetGapLapData(OptionType As Integer)
    Dim WhereX As String
    If OptionType < 3 Then WhereX = "WHERE (((UnitExitOption.OptionTypeID)=" & OptionType & "));" Else WhereX = "WHERE (((UnitExitOption.OptionTypeID) > 2));"
    SQLz = "SELECT UnitID, GapLapMsg INTO ExitGapLapTemp " & _
           "FROM EXITOptions_WorkTable INNER JOIN ExitOptionsGapLapStatus ON EXITOptions_WorkTable.UnitID = ExitOptionsGapLapStatus.AssetID;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "UPDATE ExitGapLapTemp INNER JOIN UnitExitOption ON ExitGapLapTemp.UnitID = UnitExitOption.UnitID " & _
           "SET UnitExitOption.EarlyGapOverLapMsg = [ExitGapLapTemp].[GapLapMsg] " & WhereX
    DoCmd.RunSQL (SQLz)
    DoCmd.DeleteObject acTable, "ExitGapLapTemp"
    DoCmd.SetWarnings True
End Sub

Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From EXITOptions_WorkTable;")
    DoCmd.SetWarnings True
End Sub
Private Sub ReQueryAll()
    Me.ExitOptions_subform_Exchange.Requery
    Me.ExitOptions_subform_Extend.Requery
    Me.ExitOptions_subform_Terminate.Requery

End Sub

Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, Client, MLOrig From MstrLease WHERE Client =" & Me.ClientList.Column(0) & " Order BY MLOrig;"
    Me.LeaseList.RowSource = SQLz
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.BankList.RowSource = ""
    Me.SchID = ""
    Me.Units_List.RowSource = ""

    Me.SchNo = Me.ClientList.Column(1)
    Me.ClientGroupID = Me.ClientList.Column(4)
    Me.ExitOptions_subform_Exchange.Visible = False
    Me.ExitOptions_subform_Extend.Visible = False
    Me.ExitOptions_subform_Terminate.Visible = False
    ' Call LeaseList_Click
    Call GetUnitCounts
End Sub

Public Sub Leaselist_Click()
    SQLz = "SELECT * From EXITSchedule WHERE MLID=" & Me.LeaseList.Column(0) & " Order By SchID;"
    Me.ScheduleList.RowSource = SQLz
    Me.GroupList.RowSource = ""
    Me.BankList.RowSource = ""
    Me.SchID = ""
    Me.SchNo = Me.ClientList.Column(1) & " MLA#: " & LeaseList.Column(1)
    Me.Units_List.RowSource = ""
    Me.ExitOptions_subform_Exchange.Visible = False
    Me.ExitOptions_subform_Extend.Visible = False
    Me.ExitOptions_subform_Terminate.Visible = False
    Call GetUnitCounts
End Sub

Public Sub GetUnitCounts()
    Dim UnitX, UnitY, UnitZ As String
    If ([Units_List].[ListCount] - 1) = 1 Then UnitX = " Select " Else UnitX = " Select All "
    If ([Units_List].[ListCount] - 1) = 1 Then UnitY = "Unit " Else UnitY = "Units "
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " Else UnitZ = "Units "

    Me.GetAll.caption = UnitX & vbNewLine & ChrW(9660) & ([Units_List].ListCount - 1) & " " & UnitY & ChrW(9660)

    If ([Units_List].[ListCount] - 1) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
    If [Units_List].[ItemsSelected].[Count] > 0 Then
        If Me.LeaseList.Column(3) = "FA" Then
            Me.ExchangeSelected.Visible = True
            Me.ExtendSelected.Visible = True
            Me.TerminateSelected.Visible = True
            Me.ExchangeSelected.caption = "ExchangeIT (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ & ChrW(9658)
            Me.ExtendSelected.caption = "ExtendIT (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ & ChrW(9658)
            Me.TerminateSelected.caption = "Terminate (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ & ChrW(9658)
            Me.BuyITSelection.Visible = False
        Else
            Me.ExchangeSelected.Visible = False
            Me.ExtendSelected.Visible = False
            Me.TerminateSelected.Visible = False
            Me.BuyITSelection.Visible = True
            Me.BuyITSelection.caption = "BuyIT (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ & ChrW(9658)
        End If
    Else
        Me.ExchangeSelected.Visible = False
        Me.ExtendSelected.Visible = False
        Me.TerminateSelected.Visible = False
        Me.BuyITSelection.Visible = False
    End If
End Sub
Public Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.SchNo = Me.ClientList.Column(1) & " MLA#: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(2)
    Call SetSQLu
    If FAUserAccess < 80 Then
        If Me.LeaseList.Column(3) = "FA" Then
        SQLz = SQLu & "Where SchID=" & Me.SchID & " AND MLOrig='FA' Order BY AssetID;"
        End If
    Else
        SQLz = SQLu & "Where SchID=" & Me.SchID & " Order BY AssetID;"
    End If
    Me.Units_List.RowSource = SQLz
    Me.ExitOptions_subform_Exchange.Visible = False
    Me.ExitOptions_subform_Extend.Visible = False
    Me.ExitOptions_subform_Terminate.Visible = False

    SQLz = "SELECT GroupId, UnitGroup, Count(AssetID) AS Units, SchID " & _
           "From ExitOptionsGroup_Local " & _
           "GROUP BY GroupId, UnitGroup, SchID " & _
           "HAVING SchID=" & Me.ScheduleList & " Order BY UnitGroup; "
    Me.GroupList.RowSource = SQLz

    SQLz = "SELECT BankId, 'Assignee Bank: ' & Assignee as Bank, Sum(AssigneeBanks.Units) AS Units, MLID, SchID From AssigneeBanks " & _
           "GROUP BY BankId, Assignee, MLID, SchID, SchID " & _
           "HAVING AssigneeBanks.SchID =" & Me.ScheduleList & ";"
    Me.BankList.RowSource = SQLz
    Me.SearchAsset = ""
    Me.SearchtDesc = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Call ClearList("GroupList")
    Call GetUnitCounts
End Sub
Public Sub GroupList_Click()
    Call SetSQLu
    If FAUserAccess < 80 Then
        If Me.LeaseList.Column(3) = "FA" Then SQLz = SQLu & "Where GroupId=" & GroupList & " AND MLORIG='FA' Order BY AssetID;"
    Else
        SQLz = SQLu & "Where GroupId=" & GroupList & " Order BY AssetID;"
    End If
    Me.Units_List.RowSource = SQLz
    Me.SchNo = Me.ClientList.Column(1) & " MLA#: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(2) & " Grp: " & GroupList.Column(1)
    Me.SearchAsset = ""
    Me.SearchtDesc = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Call GetUnitCounts
End Sub
Private Sub EOLYear_Click()
    Call SetSQLu
    SQLz = SQLu & "Where EOLYear=" & Me.EOLYear & " Order by ClientID and EOLDate Desc;"
    Me.SchNo = "All Units with end lease dates in " & Me.EOLYear
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
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
Private Sub SearchAsset_Change()
  ' Me.SearchAsset = ""
    Me.SearchtDesc = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Call SetSQLu
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SQLz = " Where GroupID=" & Me.GroupList & " "
    Else
        SQLz = " Where SchID=" & Me.SchID & " "
    End If
    SQLz = SQLz & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLu & SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchUnitNum_Change()
    Me.SearchAsset = ""
    Me.SearchtDesc = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
  ' Me.SearchUnitNum = ""
    Call SetSQLu
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SQLz = " Where GroupID=" & Me.GroupList & " "
    Else
        SQLz = " Where SchID=" & Me.SchID & " "
    End If
    SQLz = SQLz & " AND UnitNum Like '*" & Me.SearchUnitNum.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLu & SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchtDesc_Change()
   'Me.SearchtDesc = ""
    Me.SearchAsset = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Call SetSQLu
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SQLz = " Where GroupID=" & Me.GroupList & " "
    Else
        SQLz = " Where SchID=" & Me.SchID & " "
    End If
    SQLz = SQLz & " AND UnitDesc Like '*" & Me.SearchtDesc.Text & "*' Order BY UnitDesc;"
    Me.Units_List.RowSource = SQLu & SQLz
    Call GetUnitCounts
End Sub

Private Sub SearchVIN_Change()
    Me.SearchtDesc = ""
    Me.SearchAsset = ""
   'Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Call SetSQLu
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SQLz = " Where GroupID=" & Me.GroupList & " "
    Else
        SQLz = " Where SchID=" & Me.SchID & " "
    End If
    SQLz = SQLz & " AND VIN Like '*" & Me.SearchVin.Text & "*' Order BY VIN;"
    Me.Units_List.RowSource = SQLu & SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Me.SearchtDesc = ""
    Me.SearchAsset = ""
    Me.SearchVin = ""
   'Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Call SetSQLu
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SQLz = " Where GroupID=" & Me.GroupList & " "
    Else
        SQLz = " Where SchID=" & Me.SchID & " "
    End If
    SQLz = SQLz & " AND uStatus Like '*" & Me.SearchStatus.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLu & SQLz
    Call GetUnitCounts
End Sub


Private Sub ClearLabel_Click()
    Me.SearchtDesc = ""
    Me.SearchAsset = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchUnitNum = ""
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub

Public Sub RemoteSelectUnit(ByVal UnitIDx As Long)
    Dim lngX As Long
    With Me.Units_List
        For lngX = Abs(.ColumnHeads) To (.ListCount - 1)
            If .ItemData(lngX) Like UnitIDx Then
                .Selected(lngX) = True
            End If
        Next
    End With
    UnitIDz = UnitIDx

End Sub
Public Sub RemoteSelectSingleUnit(UnitIDx As Long)
    Dim lngX As Long
    With Me.Units_List
        For lngX = Abs(.ColumnHeads) To (.ListCount - 1)
            .Selected(lngX) = False
        Next
        For lngX = Abs(.ColumnHeads) To (.ListCount - 1)
            '---------> Use this line to test the row's bound-column value:
            If .ItemData(lngX) Like UnitIDx Then
                '---------> Use this line to test a particular column:
                'If .Column(1, lngX) Like "ABC*" Then
                .Selected(lngX) = True
            End If
        Next
    End With
End Sub

Public Sub Units_List_Click()
    Dim ClientIDx, MLAx, Schx, Grpx As Integer
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub

    On Error GoTo FixMe
    ClientIDx = DLookup("ClientID", "vw_SixKeys", "AssetID=" & Me.Units_List.Column(0))
    Me.ClientList = ClientIDx
    SQLz = "SELECT MLID, MLNo, Client, MLOrig From MstrLease WHERE Client =" & Me.ClientList.Column(0) & " Order BY MLOrig;"
    Me.LeaseList.RowSource = SQLz

    MLAx = DLookup("MLID", "vw_SixKeys", "AssetID=" & Me.Units_List.Column(0))
    SQLz = "SELECT * From EXITSchedule WHERE MLID=" & MLAx & " Order By SchID;"
    Me.ScheduleList.RowSource = SQLz
    Me.LeaseList = MLAx

    Schx = DLookup("SchID", "vw_SixKeys", "AssetID=" & Me.Units_List.Column(0))
    Me.ScheduleList = Schx
    SQLz = "SELECT GroupId, UnitGroup, Count(AssetID) AS Units, SchID From ExitOptionsGroup_Local " & _
           "GROUP BY GroupId, UnitGroup, SchID HAVING SchID=" & Me.ScheduleList & " Order BY UnitGroup; "
    Me.GroupList.RowSource = SQLz
    Me.SchID = Schx
    Grpx = DLookup("SGrpId", "vw_SixKeys", "AssetID=" & Me.Units_List.Column(0))
    Me.GroupList = Grpx
    Me.Units_List.ControlTipText = Nz(Units_List.Column(3), "test")
    Me.SchNo = Me.ClientList.Column(1) & " MLA#: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(2) & " Grp: " & GroupList.Column(1)
    Call GetUnitCounts
'    Call GroupList_Click


    'Me.GroupList = DLookup("SGrpId", "vw_SixKeys", "SchID=" & Me.ScheduleList.Column(0) & "")
    '    On Error GoTo Fixme
    If ActionType = "History" Then
  '      SQLz = "Select * From ExitOptionsQryAll Where UnitId =" & Me.Units_List.Column(0) & " ORDER BY UnitId, OptionTypeDesc, [Start-Purposed];"
  '      Me.ExitOptions_subform_History.Form.RecordSource = SQLz

    End If

FixMe:
    On Error GoTo 0
    ' traps errors
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
' On Error GoTo Fixme
' Set RST = Me.ExitOptions_subform_Rents.Form.RecordsetClone
' RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
' Me.ExitOptions_subform_Rents.Form.Bookmark = RST.Bookmark
' Set rs = Nothing
FixMe:
End Sub
