Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ModifyExch_Click()
    Dim SchIDz As Integer
    SchIDz = Me.SchID
    If Me.Sts <> "A" And Me.Sts <> "TR" Then
        MsgBox "An EXIT Option may only applied to an A or TR unit." & vbNewLine & "  If you need to apply an EXIT Option," & vbNewLine & "  Please update the status of the Unit." & vbNewLine & "    or" & vbNewLine & "  Create a Ticket with IT", , "EXIT Option Create Warning"
        Exit Sub
    End If
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = Me.GroupID
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.RemoteSelectUnit (Me.UnitID)
    Form_EXITOptions.ExchangeSelected_Click
End Sub


Private Sub ModifyEXTD_Click()
    Dim SchIDz As Integer
    SchIDz = Me.SchID
    If Me.Sts <> "A" And Me.Sts <> "TR" Then
        MsgBox "An EXIT Option may only applied to an A or TR unit." & vbNewLine & "  If you need to apply an EXIT Option," & vbNewLine & "  Please update the status of the Unit." & vbNewLine & "    or" & vbNewLine & "  Create a Ticket with IT", , "EXIT Option Create Warning"
        Exit Sub
    End If
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = Me.GroupID
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.RemoteSelectUnit (Me.UnitID)
    Form_EXITOptions.ExtendSelected_Click
End Sub

Private Sub ModifyTERM_Click()
    Dim SchIDz As Integer
    SchIDz = Me.SchID
    If Me.Sts <> "A" And Me.Sts <> "TR" Then
        MsgBox "An EXIT Option may only applied to an A or TR unit." & vbNewLine & "  If you need to apply an EXIT Option," & vbNewLine & "  Please update the status of the Unit." & vbNewLine & "    or" & vbNewLine & "  Create a Ticket with IT", , "EXIT Option Create Warning"
        Exit Sub
    End If
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = Me.GroupID
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.RemoteSelectUnit (Me.UnitID)
    Form_EXITOptions.TerminateSelected_Click
End Sub

Private Sub ModifyHIST_Click()
    Dim SchIDz As Integer
    SchIDz = Me.SchID
    If DCount("UnitID", "UnitExitOption", "UnitID=" & Me.UnitID) < 1 Then
        MsgBox "AssetID: " & Me.UnitID & " does not have an EXIT Option applied nor any EXIT Option history." & vbNewLine & "To save time, the Exit Options History Form will NOT be loaded." & vbNewLine & vbNewLine & "If the Unit should have EXIT Option history, please create a Ticket with IT.", , "Unit has No History"
        Exit Sub
    End If

    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = Me.GroupID
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.RemoteSelectUnit (Me.UnitID)
    '    Forms!ExitOptions.Units_List.Column(0) = Me.UnitID
    Form_EXITOptions.Units_List_Click
    Form_EXITOptions.UnitHistory_click
    'Form_EXITOptions.History_Click
    'Call UnitID_DblClick
End Sub
Private Sub NewUnit_DblClick(Cancel As Integer)
    If Me.NewUnit & "test" = "test" Then Exit Sub
    Call OpenUnitsForm(Me.NewUnit)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.NewUnit
    Forms!Unitsform.UnitsTabControl.Value = 8
    Me.unitunitnum.SetFocus
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
' DoCmd.close acForm, "UnitsForm"
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID
    Forms!Unitsform.UnitsTabControl.Value = 8
    Me.unitunitnum.SetFocus
End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
    Call UnitID_DblClick(1)
End Sub
