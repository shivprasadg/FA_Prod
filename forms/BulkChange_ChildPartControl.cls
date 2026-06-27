Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx As String
Dim SQLz As String
Dim Response As String


Private Sub ChildList_Click()
  '  SQLz = "Select DISTINCTROW * From Bulk_Update_SelectedUnits_Child Where UID ='" & Me.ChildList & " ';"
  '  Me.Form.RecordSource = SQLz
     Me.UnitCountLabel.BackColor = RGB(192, 80, 77)
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DeleteChildPart_Click()
    Dim Response As String
    Response = MsgBox("Are you sure you want to Delete Child Part: '" & Me.ChildList.Column(0) & "'?" & vbNewLine & "(this will delete that part from all " & Me.UnitCount & " units)", vbQuestion + vbYesNo + vbCritical, "MDK!!")
    If Response = vbYes Then
        SQLz = "DELETE DISTINCTROW Bulk_Update_SelectedUnits_Child.UID, ChildParts.*, * " & _
               "FROM Bulk_Update_SelectedUnits_Child INNER JOIN ChildParts ON Bulk_Update_SelectedUnits_Child.ChildPartID = ChildParts.ChildPartID " & _
               "WHERE Bulk_Update_SelectedUnits_Child.UID ='" & Me.ChildList.Column(3) & "';"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        ModMade = -1
        Me.ChildList.Requery
        DoCmd.SetWarnings True
        Me.UnitCountLabel.BackColor = RGB(0, 50, 102)
    End If
End Sub


Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub

Private Sub InsertChildParts_Click()
    Response = MsgBox("Are you sure you want to Insert this Child Part:" _
                      & vbNewLine & Me.MakeCbo & " " & Me.ClassCbo & " " & Me.TypeCbo.Column(0) & " ?" _
                      & vbNewLine & "(this will be added to " & Me.UnitCount & " units)", vbQuestion + vbYesNo + vbCritical, "MDK!!")
    If Response = vbYes Then
       Call MassInsertParts
    End If
End Sub
Public Sub InsertSchChildParts_Click()
    Response = MsgBox("Are you sure you want to Add this Child Part:" _
                      & vbNewLine & Me.MakeCbo & " " & Me.ClassCbo & " " & Me.TypeCbo.Column(0) & " ?" _
                      & vbNewLine & "This will be added ALL Units in this SCHEDULE!!", vbQuestion + vbYesNo + vbCritical, "SCH Insert")
    If Response = vbYes Then
    Call MassInsertParts(Nz(Me.Parent.Form.ScheduleList.Column(0), 0))
    End If
End Sub

Public Function MassInsertParts(Optional SchIDx As Long)
    If Me.UnitCount < 1 Then
        MsgBox "You must select a range of Parent Units", , "No Units Selected"
        Exit Function
    End If

    Dim Msgx As String
    Dim UnitIDx As String

    Me.Parent.BulkChange_ChildParts.Form.Refresh
    Msgx = ""
    If Me.ChildVendorCbo & "test" = "test" Then Msgx = Msgx & " Vendor |"
    If Me.MakeCbo & "test" = "test" Then Msgx = Msgx & " Make |"
    If Me.ClassCbo & "test" = "test" Then Msgx = Msgx & " Model |"
    If Me.TypeCbo & "test" = "test" Then Msgx = Msgx & " Type |"
    If Me.CostCurr & "test" = "test" Then Msgx = Msgx & " Cost |"
    If Msgx <> "" Then
        MsgBox "You're missing the following information:" & vbNewLine & "| " & Msgx, , "Need more info"
        Exit Function
    End If
    ' Insert the ChildPart
    SQLz = "INSERT INTO ChildParts ( "
    SQLz = SQLz & "UnitREF"
    SQLz = SQLz & ", ChildMake"
    SQLz = SQLz & ", ChildClass"
    SQLz = SQLz & ", ChildType"
    SQLz = SQLz & ", ChildVendorREF"
    SQLz = SQLz & ", childFinalCost"
    SQLz = SQLz & ", PartAddedDate"
    'SQLz = SQLz & ", PartAddedBy
    'Add CDesc if CDescCbo is filled
    If Me.cDescCbo & "test" <> "test" Then SQLz = SQLz & ", cDesc"

    SQLz = SQLz & ") SELECT UnitID"
    SQLz = SQLz & ",'" & Me.MakeCbo & "'"             'AS ChildMake"
    SQLz = SQLz & ",'" & Me.ClassCbo & "'"            'AS ChildClass"
    SQLz = SQLz & ",'" & Me.TypeCbo.Column(0) & "'"   'AS ChildType"
    SQLz = SQLz & ", " & Me.ChildVendorCbo & " "      'AS ChildVendorREF"
    SQLz = SQLz & ", " & Me.CostCurr & " "            'AS childFinalCost"
    SQLz = SQLz & ", #" & Date & "# "                 'AS PartAddedDate"
    'SQLz = SQLz & ", '" & FAusername & "' "          'AS PartAddedBy"
    ' Add CDesc if CDescCbo is filled
    If Me.cDescCbo & "test" <> "test" Then SQLz = SQLz & ", '" & Me.cDescCbo & "' " 'AS cDesc"

    If SchIDx > 0 Then
        SQLz = SQLz & " FROM vw_SixKeys WHERE SchID= " & SchIDx & ";"

        'Delete the SAME Child Part if it Exist in the Same Sch
        SQLx = "DELETE From ChildParts"
        SQLx = SQLx & " WHERE ChildMake='" & Me.MakeCbo & "' "
        SQLx = SQLx & " AND ChildClass = '" & Me.ClassCbo & "' "
        SQLx = SQLx & " AND ChildType='" & Me.TypeCbo & "' "
        SQLx = SQLx & " AND UnitRef in (Select UnitID FROM vw_SixKeys WHERE SchID= " & SchIDx & ")"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLx)
        DoCmd.SetWarnings True
    Else
        SQLz = SQLz & " FROM Bulk_Update_WorkTable;"
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    'Check for Interim Funding
    Call CheckInterimFunded(Me.MakeCbo & " " & Me.ClassCbo & " " & Me.TypeCbo.Column(0) & " " & Format(Me.CostCurr, "$#,#.#0") & "")
    ModMade = -1
    Call ClearBoxes
    Me.Parent.BulkChange_ChildParts.Form.Refresh
    Me.ChildList.Requery
    Me.SisterPArts.Requery
End Function

Private Sub ClassCbo_Click()
    If Me.ClassCbo.Column(1) <> Nz(Me.TypeCbo.Column(1), 0) Then
       Me.TypeCbo = ""
       Me.cDescCbo = ""
    End If
    SQLz = "SELECT cPartType, cPartClassID " & _
           "FROM Lookup_PartTypes " & _
           "WHERE cPartClassID=" & Me.ClassCbo.Column(1) & " " & _
           "Order By cPartType;"
    Me.TypeCbo.RowSource = SQLz
End Sub
Private Sub TypeCbo_GotFocus()
    If Me.ClassCbo & "test" = "test" Then
        MsgBox "You MUST Select a Class", , "No can do, without a Class"
        Me.ClassCbo.SetFocus
    End If
End Sub

Private Sub TypeCbo_Click()
'    Me.TypeCbo = Nz(DLookup("ChildPartTypeID", "LookUp_ChildPartType", "ChildPartType='" & Me.ModelCbo.Column(2) & "'"), 1)
End Sub
