Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLset As String
Public Committ As Variant
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub



Private Sub DeleteMe_Click()
    Committ = MsgBox("You're about set the Delete" & vbNewLine & "  Inspection Num: " & Me.InspNum & vbNewLine & "  Dated: " & Me.iDate & vbNewLine & vbNewLine & "Are You Sure?", vbYesNo + vbCritical, "Vendor Change?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Delete From UnitInspections Where InspnID =" & Me.InspnID
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "Delete From UnitInsp_WorkTable Where InspnID =" & Me.InspnID
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
     Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.UnitID
    Dim RST        As DAO.Recordset
    SQLz = "Select ClientShortname As Client, Schedule, UnitGroup From vw_SixKeys WHERE AssetID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.Parent.BuyNo = "Client: " & RST!Client & " Sch|Grp: " & RST!Schedule & "|" & RST!UnitGroup
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Sub

Private Sub GetDekraData1_Click()
    OpenInspectionFileDialog (1)
    Me.Form.Refresh
End Sub
Private Sub GetDekraData2_Click()
    OpenInspectionFileDialog (2)
    Me.Form.Refresh
End Sub


'Private Sub InspectionLinkLocal_AfterUpdate()
'    Me.InspectionLinkLocal = "Local#" & Me.InspectionLinkLocal & "#"
'    If IsNull(Me.InspectionLinkLocal) Or Me.InspectionLinkLocal Like "\\*" Or Me.InspectionLinkLocal Like "P:*" Then
'    Else
'        Dim Response As String
'        Response = MsgBox("This document might not be accessible" & vbNewLine & "by other users, over the network." & vbNewLine & vbNewLine & "Do you still want to link it?", vbQuestion + vbYesNo + vbCritical, "Not A Network Location")
'        If Response = vbNo Then
'            On Error GoTo Fixme
'            Me.InspectionLinkLocal = Null
'        End If
'    End If
'    Exit Sub
'Fixme:
'    Me.InspectionLinkLocal = Null
'End Sub
Private Sub DeleteWebLink_Click()
    Me.InspectionLinkWeb = Null
End Sub

Private Sub InspectionLinkWeb_AfterUpdate()
    Me.InspectionLinkWeb = "Web#" & Me.InspectionLinkWeb & "#"
    If IsNull(Me.InspectionLinkWeb) Or Me.InspectionLinkWeb Like "*//*" Or InspectionLinkWeb Like "*www*" Or InspectionLinkWeb Like "*htt*" Then
    Else
        Dim Response As String
        Response = MsgBox("This document might not be accessible" & vbNewLine & "by other users, over the network." & vbNewLine & vbNewLine & "Do you still want to link it?", vbQuestion + vbYesNo + vbCritical, "Not A Web Location")
        If Response = vbNo Then
            On Error GoTo FixMe
            Me.InspectionLinkWeb = Null
        End If
    End If
    Exit Sub
FixMe:
    Me.InspectionLinkWeb = Null
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Committ As String
    If Me(FieldBoX).Value & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update-Change confrim " & TragetX & "?")
        If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
End Sub

Private Sub iPassedDate_AfterUpdate()
    If IsDate(Me.iPassedDate) Then Me.iResultID = 1 Else Me.iResultID = 3
End Sub

Private Sub iPassedDateSubmit_Click()
    Committ = MsgBox("You're about set the Passed date to '" & iPassedDateCBO & "'" & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Vendor Change?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update UnitInspections "
    If IsDate(iPassedDateSubmit) Then
        SQLz = SQLz & " Set iPassedDate = #" & Me.iPassedDateCBO & "# "
    Else
        SQLz = SQLz & " Set iPassedDate = Null "
    End If
    SQLz = SQLz & " Where InspnID in (Select InspnID From UnitInsp_WorkTable)"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub iResultID_AfterUpdate()
   Call AutoUpdateEOL_Statuses(True)
End Sub

Private Sub iResultIDSubmit_Click()
    Committ = MsgBox("You're about set the Results to '" & iResultIDCbo.Column(1) & "'" & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Vendor Change?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update UnitInspections Set iResultID=" & Me.iResultIDCbo & " Where InspnID in (Select InspnID From UnitInsp_WorkTable)"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call AutoUpdateEOL_Statuses(True)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub iTypeIDSubmit_Click()
    Committ = MsgBox("You're about set the Type to '" & iTypeIDCbo.Column(1) & "'" & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Vendor Change?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update UnitInspections Set iTypeID=" & Me.iTypeIDCbo & " Where InspnID in (Select InspnID From UnitInsp_WorkTable)"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub iVendorIDSubmit_Click()
    Committ = MsgBox("You're about set the Vendor to '" & iVendorCBO.Column(1) & "'" & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Vendor Change?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update UnitInspections Set iVendorId=" & Me.iVendorCBO & " Where InspnID in (Select InspnID From UnitInsp_WorkTable)"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|R,HY|"
    SendFormData2Excel Me, "Inspections", "Inspection Info", True, XLS, 0
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    Forms!Unitsform.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub
