Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExcelReport_Click()
    Dim XLS, Header As String
    Dim x As Integer
    Header = "Unit List " & Me.Parent.SchNo & " " & Date
    LScript = "|"
    For x = 1 To 19
        XLS = XLS & "FJ,ZZ|"
    Next x
    For x = 1 To 52
        XLS = XLS & "DB,ZZ|"
    Next x
    XLS = XLS & "CK,ZZ|CG,ZZ|CF,ZZ|CE,ZZ|AY,ZZ|AR,ZZ|AJ,ZZ|AI,ZZ|AH,ZZ|AG,ZZ|AF,ZZ|AI,ZZ|AI,ZZ|AI,ZZ|AF,ZZ|AF,ZZ|AF,ZZ|AF,ZZ|"
    For x = 1 To 11
        XLS = XLS & "V,ZZ|"
    Next x
    XLS = XLS & "F,TX|"
    SendReport2Excel Me.Form.RecordSource, "Fleet List", Header, True, XLS, Me.UnitCount, True
End Sub

Private Sub Form_AfterUpdate()
    Me.Parent.ModMade = -1
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.Parent.ModMade = -1
End Sub

Private Sub Form_Load()
    Me.SubTypeCbo.RowSource = ""
    Me.ConfigCBO.RowSource = ""
    Me.UOMCbo.RowSource = ""
End Sub
Private Sub Form_Close()
    If Me.GroupRef & "test" <> "test" Then
        If DMax("Status", "vw_SixKeys", "SGrpID=" & Me.GroupRef) = "A" Then Call AllUnitsInGroup_AcceptedMessage(Me.GroupRef)
    End If
End Sub
Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub
Private Sub TypeCbo_Click()
    Me.SubTypeCbo = Null
    Me.ConfigCBO = Null
    SQLz = "SELECT UnitSubType, UnitSubTypeDesc, UnitSubTypeID, UnitTypeID " & _
        "FROM LookUp_UnitSubTypes Where UnitTypeID =" & Me.TypeCbo.Column(2) & " Order By UnitSubType;"
    Me.SubTypeCbo.RowSource = SQLz
    Call SetUOM(Nz(Me.SubTypeCbo, ""), Nz(Me.TypeCbo, ""), "UOMCbo")
End Sub
Private Sub SubTypeCbo_Click()
    Me.ConfigCBO = Null
    SQLz = "SELECT UnitConfig, UnitConfigDesc, UnitConfigID, UnitSubTypeID " & _
        "FROM LookUp_UnitConfiguration Where UnitSubTypeID =" & Me.SubTypeCbo.Column(2) & " Order By UnitConfig;"
    Me.ConfigCBO.RowSource = SQLz
    Me.TypeCbo = DLookup("UnitType", "LookUp_UnitTypes", "UnitTypeID=" & Me.SubTypeCbo.Column(3))
    Call SetUOM(Nz(Me.SubTypeCbo, ""), Nz(Me.TypeCbo, ""), "UOMCbo")
End Sub

Private Sub MakeCbo_AfterUpdate()
    Me.ModelCbo = ""
    Me.ModelCbo.Requery
End Sub

Private Sub MakeCbo_Click()
    Me.ModelCbo = ""
    Me.ModelCbo.Requery
End Sub

Private Sub ModelCbo_GotFocus()
    Me.ModelCbo.Requery
    If Me.MakeCbo & "test" = "test" Then
        MsgBox "You MUST Select a Make", , "No can do, without a Make"
        Me.MakeCbo.SetFocus
    End If
End Sub

Private Sub AxleSubmit_Click()
    On Error GoTo FixMe
    If Me.YearCbo & "test" = "test" Then Call UpdateMassUnits("UnitYr", Null)
    Call UpdateMassUnits("UnitYr", Me.YearCbo)
    ModMade = -1
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub AxlePropertySubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitAxleProperty", Me.AxlePropertyCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub DescriptorSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitDescriptor", Me.DescriptorCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub MakeSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitMake", Me.MakeCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub ModelSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitModel", Me.ModelCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub SubTypeSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitSubType", Me.SubTypeCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub TypeSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitType", Me.TypeCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitConfigSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitConfiguration", Me.ConfigCBO)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitUOMSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitUOM", Me.UOMCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub YearSubmit_Click()
    If Me.YearCbo & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateMassUnits("UnitYr", Me.YearCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.unitunitnum.SetFocus
End Sub

Private Sub UnitMake_AfterUpdate()
    Me.UnitModel.Requery
End Sub

Private Sub UnitStatus_BeforeUpdate(Cancel As Integer)
    If Me.UnitStatus Like "*R*" Or Me.UnitStatus Like "*P*" Or Me.UnitStatus Like "*S*" Then
        MsgBox "You cannot use a Remarketing Unit Status here" & vbNewLine & "Use the Remarketing Mass Mode form, or" & vbNewLine & "you may use the Units form directly", vbCritical, "Danger"
        Me.Undo
        Exit Sub
    End If
End Sub

Private Sub UnitStatusSubmit_Click()
    On Error GoTo FixMe
    If Me.UnitStatusCbo Like "*P*" Or Me.UnitStatusCbo Like "*R*" Or Me.UnitStatusCbo Like "*S*" Or Me.UnitStatusCbo Like "*E*" Then
        MsgBox "You cannot use a Remarketing Unit Status here" & vbNewLine & "Use the Remarketing Mass Mode form, or" _
              & vbNewLine & "you may use the Units form directly", vbCritical, "Danger Mouse"
        Exit Sub
    End If
    On Error GoTo FixMe
    If Me.UnitStatusCbo = "A" Then
       Call UpdateMassUnits("Units.UnitStatus", Me.UnitStatusCbo & "', Units.UnitDocAccpt=-1 ", "WHERE (IsDate([UnitCOABack]) OR IsDate([UnitCofACopy])) AND UnitStatus <> 'A' ")
    Else
       Call UpdateMassUnits("Units.UnitStatus", Me.UnitStatusCbo)
    End If
    ModMade = -1
    If Me.Parent.LeaseList.Column(2) = "FA" And AllUnitsInGroup_Accepted(Me.GroupRef) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.GroupRef)
    Me.Form.Refresh
    If Me.UnitStatusCbo = "A" Then Exit Sub
FixMe:
End Sub

Private Sub UnitType_Click()
    Me.UnitSubType.Requery
    Call SetUOM(Nz(Me.UnitSubType, ""), Nz(Me.UnitType, ""), "UnitUom")
End Sub
Private Sub UnitSubType_Click()
    Me.UnitConfiguration.Requery
    Call SetUOM(Nz(Me.UnitSubType, ""), Nz(Me.UnitType, ""), "UnitUom")
End Sub
Private Sub UnitUOM_GotFocus()
    Call SetUOM(Nz(Me.UnitSubType, ""), Nz(Me.UnitType, ""), "UnitUom")
End Sub
Private Sub UOMCbo_GotFocus()
    Call SetUOM(Nz(Me.SubTypeCbo, ""), Nz(Me.TypeCbo, ""), "UOMCbo")
End Sub
Public Sub SetUOM(UnitSubTypeX, UnitTypeX, ControlNamex As String)
    ' Set UOM Control Tip and ComboBox
    Dim UomX, ix, RangeLow, RangeHigh, StepX As Integer
    Dim ListX, MeasureType As String
    Me(ControlNamex).ColumnWidths = "0;0.8"
    If UnitTypeX = "Trailer" Then UomX = 30
    If UnitSubTypeX = "Tank" Then UomX = 37
    If UnitSubTypeX = "Sleeper" Then UomX = 12
    If UnitSubTypeX = "Daycab" Then UomX = 12
    If UnitTypeX = "Van" Then UomX = 22
    If UnitSubTypeX = "Auto Hauler" Then UomX = 36
    If UnitTypeX = "Straight Truck" Then UomX = 20
    If UnitSubTypeX = "Tank" And UnitTypeX = "Straight Truck" Then UomX = 27
    If UomX = 27 Or UomX = 37 Then StepX = 500 Else StepX = 1
    If UomX <> 0 Then
        RangeLow = DMin("UOMRangeLow", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        RangeHigh = DMax("UOMRangeHigh", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        MeasureType = DLookup("UOMPrintLabel", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        Me(ControlNamex).ControlTipText = "A " & SubUnitTypeX & " " & UnitTypeX & "'s UOM Range is: " & DLookup("UOMLabel", "vw_UnitType_Matrix", "UOMGroupID = " & UomX)

        For ix = RangeLow To RangeHigh Step StepX
            ListX = ListX & ix & ";" & ix & MeasureType & ";"
        Next ix

        Me(ControlNamex).RowSource = ListX
        Me(ControlNamex).Enabled = True
    Else
        Me(ControlNamex).RowSource = ""
        Me(ControlNamex).ControlTipText = "A " & UnitSubTypeX & " " & UnitTypeX & " doesn't have a UOM"
    End If
End Sub
