Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_AfterUpdate()
    Me.Parent.ModMade = -1
End Sub

Private Sub Form_Current()
    On Error GoTo whoops
    Me.txtCurrent = Me.UnitID
    If Nz(Me.UnitIFAmt, 0) > 0 And FAUserAccess < 99 Then
        Me.UnitFinalCost.Locked = True
        Me.CostSubmit.Enabled = False
    Else
        Me.UnitFinalCost.Locked = False
        Me.CostSubmit.Enabled = True
    End If

whoops:
    Exit Sub
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
End Sub
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
    ' Me.DrawDownCbo.Requery
    ' Me.DrawDownId.Requery
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub
Private Sub BuildDateSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitBuildDate", Me.BuildDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub EstDeliveryDateSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitEstDelDate", Me.EstDeliveryDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub ActualDelDateSubmit_Click()
    On Error GoTo FixMe
    Dim sWhere As String
    If Me.ClientGroupID = 187 Then sWhere = "WHERE IsNull(Units.UnitAcceptDate) AND NOT UnitActualDelDate Is Null" Else sWhere = ""
    Call UpdateMassUnits("Units.UnitActualDelDate", Me.ActualDelDate, sWhere)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub AcceptedDateSubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.AcceptedDate & "test" <> "test" Then
       ExtraFields = ", Units.UnitDocAccpt = -1, Units.unitstatus = 'A' "
       Call UpdateMassUnits("Units.UnitAcceptDate", Me.AcceptedDate, "WHERE IsNull(Units.UnitAcceptDate)", ExtraFields)
    End If
FixMe:
    Me.Form.Refresh
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub

Private Sub SalesResidSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitSalesResid", Me.SalesResidCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub StripCurr_AfterUpdate()
    If Me.PDRentCurr & "test" = "test" Then
        Me.PDRentCurr = Round(Me.StripCurr / 30, 2)
    End If
End Sub

Private Sub SynResidSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitSynResid", Me.SynResidCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub UnitActualDelDate_Click()
    If Me.ClientGroupID = 187 And IsNull(Me.ActualDelDate) Then Me.AcceptedDate = Me.ActualDelDate
End Sub

Private Sub UnitStatus_AfterUpdate()
Me.UnitStatus = UCase(Me.UnitStatus)
    If Not IsDate(DLookup("cofacopy", "units", "Unitid=" & Me.UnitID)) And Not IsDate(DLookup("coaback", "units", "Unitid=" & Me.UnitID)) And Me.UnitStatus = "A" Then
            MsgBox "You must have an CoA Copy Date entered", vbInformation, "CoA Date Req"
        Me.UnitStatus = "B"
        Exit Sub
    End If
    If IsDate(DLookup("cofacopy", "units", "Unitid=" & Me.UnitID)) And Me.UnitStatus = "A" Then
           Me.UnitDocAccpt = -1
    End If
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.UnitInServ.SetFocus
End Sub

Private Sub UnitPDStartDateSubmit_Click()
    On Error GoTo FixMe
    If IsDate(PDStartDateCBO) Then
       Call UpdateMassUnits("Units.UnitPDStartDate", Me.PDStartDateCBO)
       ModMade = -1
    End If
FixMe:
    Me.Form.Refresh
End Sub

Private Sub InServiceSubmit_Click()
    On Error GoTo FixMe
    If Me.InservDate & "test" = "test" Then
        Call UpdateMassUnits("Units.UnitInServ", Null)
    Else
        Call UpdateMassUnits("Units.UnitInServ", Me.InservDate)
    End If
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub CostSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitFinalCost", Me.CostCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub PDRentSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UNITPDRENT", Me.PDRentCurr)
FixMe:
    Me.Form.Refresh
End Sub
Private Sub PDSurrenderSubmit_Click()
    On Error GoTo FixMe
    If Me.PDSurrenderCurr & "test" <> "test" Then
        Call UpdateMassUnits("Units.PerDiemSurrender", Me.PDSurrenderCurr)
        ModMade = -1
        Me.Form.Refresh
    End If
FixMe:
End Sub
Private Sub StripSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UNITSRENT", Me.StripCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub SyndRentSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UNITRENT", Me.SyndRentCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub UnitTaxSubmit_Click()
    On Error GoTo FixMe
    Exit Sub
  '  SQLz = "UPDATE Units INNER JOIN BulkUpdate_UnitTaxes ON Units.UnitID = BulkUpdate_UnitTaxes.UnitID SET Units.UnitTax = BulkUpdate_UnitTaxes.UnitTaxes;"
  '  DoCmd.RunSQL (SQLz)
  '  ModMade = -1
FixMe:
  '  Me.Form.Refresh
  '  Exit Sub
End Sub
Private Sub FirstExtRentSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.FirstExtRent", Me.FirstExtCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub MTMRentSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.MTMRent", Me.MTMRentCbo)
    ModMade = -1
FixMe:
    Exit Sub
End Sub
