Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String
Dim Response   As String
Dim PType      As String

Private Sub AddNew_Click()
    If Me.Dirty Then
        MsgBox "You Cannot Add a record at this time." & vbNewLine & "You Must SAVE/UPDATE This One First", vbCritical, "Editing Data!"
        Exit Sub
        End
    End If
    Response = MsgBox("You are about to Add a Move Order." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Add Work Order?")
    If Response = vbYes Then
        Call ParkedLocationMove(Nz(DLookup("LocationParkedID", "Units", "UnitID=" & Me.mAssetID & ""), DLookup("UnitGarageLocationRef", "Units", "UnitID=" & Me.mAssetID & "")), 0, Me.mAssetID)
        If Me.mEmail & "test" <> "test" Then
            Me.Detail.Visible = True
            Me.Form.Requery
            MsgBox "Please Select the MOVE TO Location you plan to use", vbCritical, "You ain't done yet!"
            Me.AddNew.Visible = False
            Me.Parent.unitstatus = "EM"
            If mOrderNum = "TBD" Or Me.mStatus = "not ordered" Then Me.SubmitMoveOrder.Visible = True
            Me.mNewLocationID.SetFocus
        Else
            Me.Undo
            Me.AddNew.Visible = True
            Me.Detail.Visible = False
        End If
    End If
End Sub

Private Sub DeleteLocation_Click()
    Response = MsgBox("Are you sure you want DELETE the Move " & vbNewLine & _
               "From:" & Me.mOldLocationID.Column(1) & vbNewLine & _
               "To:" & Me.mNewLocationID.Column(1) & vbNewLine & _
               "Dated:" & Me.mStartDate & "?", vbCritical + vbYesNo, "Don't play with FIRE")
    If Response <> "6" Then Exit Sub
    SQLz = "Delete From UnitMoveOrder WHERE UnitMoveID =" & Me.UnitMoveID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.MoveOrderList.Requery
End Sub

Private Sub Form_Load()
    If Nz(DCount("UnitMoveID", "UnitMoveOrder", "mAssetID=" & Nz(Me.mAssetID, 0) & ""), 0) = 0 Then
        Me.AddNew.Visible = True
        Me.Detail.Visible = False
        Me.SubmitMoveOrder.Visible = False
    Else
        Me.AddNew.Visible = False
        Me.Detail.Visible = True
        If mOrderNum = "TBD" Or Me.mStatus = "not ordered" Then Me.SubmitMoveOrder.Visible = True Else Me.SubmitMoveOrder.Visible = False
    End If
    SQLz = "SELECT UnitMoveID "
    SQLz = SQLz & ", mStartDate AS [Start]"
    SQLz = SQLz & ", mEndDate AS [End]"
    SQLz = SQLz & ", PickupLocation As PickUp"
    SQLz = SQLz & ", DeliveryLocation AS Deliver"
    SQLz = SQLz & ", MStatus"
    SQLz = SQLz & ", AssetID"
    SQLz = SQLz & " FROM vw_UnitMoveDashboard"
    SQLz = SQLz & " WHERE AssetID = " & Me.mAssetID & ""
    SQLz = SQLz & " ORDER BY mStartDate;"
    Me.MoveOrderList.RowSource = SQLz
End Sub

Private Sub SaveMe_Click()
    Dim Validate As String
    If Nz(Me.mVendorID, 0) = 0 Then Validate = Validate & "   Vendor" & vbNewLine
    If Nz(Me.mOldLocationID, 0) = 0 Then Validate = Validate & "   From Location" & vbNewLine
    If Nz(Me.mNewLocationID, 0) = 0 Then Validate = Validate & "   To Location" & vbNewLine
    If Nz(Me.mStartDate, 0) = 0 Then Validate = Validate & "   Start Date" & vbNewLine
    If Nz(Me.mContact, 0) = 0 Then Validate = Validate & "   Contact" & vbNewLine
    If Nz(Me.mEmail, 0) = 0 Then Validate = Validate & "   Email" & vbNewLine
    If Nz(Me.mPhone, 0) = 0 Then Validate = Validate & "   Phone" & vbNewLine
    If Validate <> "" Then
        MsgBox "You Cannot Submit this MOVE request yet." & vbNewLine & "You're Missing:" & vbNewLine & Validate & "(nothing done)", vbCritical, "Missing Data!"
        Exit Sub
    End If
    Me.AssetID = Forms!Unitsform.Form.UnitID
    DoCmd.RunCommand acCmdSaveRecord
    Call LoadData(Forms!Unitsform.Form.UnitID, True)

End Sub



Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_RmktExpense"
End Sub


Private Sub mNewLocationID_AfterUpdate()
 Call Form_Load
End Sub

Private Sub MoveOrderList_Click()
    Me.Form.FilterOn = False

    Set RST = Me.RecordsetClone
    RST.FindFirst "UnitMoveID = " & Me.MoveOrderList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End Sub

Private Sub SubmitMoveOrder_Click()
    Dim Validate As String
    If Nz(Me.mVendorID, 0) = 0 Then Validate = Validate & "   Vendor" & vbNewLine
    If Nz(Me.mOldLocationID, 0) = 0 Then Validate = Validate & "   From Location" & vbNewLine
    If Nz(Me.mNewLocationID, 0) = 0 Then Validate = Validate & "   To Location" & vbNewLine
    If Nz(Me.mStartDate, 0) = 0 Then Validate = Validate & "   Start Date" & vbNewLine
    If Nz(Me.mContact, 0) = 0 Then Validate = Validate & "   Contact" & vbNewLine
    If Nz(Me.mEmail, 0) = 0 Then Validate = Validate & "   Email" & vbNewLine
    If Nz(Me.mPhone, 0) = 0 Then Validate = Validate & "   Phone" & vbNewLine
    If Validate <> "" Then
        MsgBox "You Cannot Submit this MOVE request yet." & vbNewLine & "You're Missing:" & vbNewLine & Validate & "(nothing done)", vbCritical, "Missing Data!"
        Exit Sub
    End If
End Sub
