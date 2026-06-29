Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub EngineSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitEngineType", Me.EngineCbo)
FixMe:
    Me.Form.Refresh
End Sub

Private Sub Form_AfterUpdate()
    Me.Parent.ModMade = -1
End Sub

Private Sub PowerSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.PowerType", Me.PowerCbo)
FixMe:
    Me.Form.Refresh
End Sub
Private Sub TransSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitTransmissionType", Me.TransCbo)
FixMe:
    Me.Form.Refresh
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    On Error GoTo FixMe    ' if the user is just sorting
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
FixMe:
End Sub
Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.unitunitnum.SetFocus
End Sub
