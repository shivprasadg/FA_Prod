Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim msg1       As String
Dim ResposeX   As VbMsgBoxResultEx
Dim addrx, Typex, TargetX As String
Private Sub Form_Load()
LocationList1.RowSource = ""
LocationList2.RowSource = ""
End Sub
Private Sub CloseMe_Click()
    DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub Form_Close()
'If Me.NewGroupList <> "" Then Call CheckATFChanges(Me.NewGroupList, "Move Units")
End Sub

Public Sub LocationList1_Click()
    Me.LocID1 = Me.LocationList1.Column(0)
    SQLz = ""
    SQLz = SQLz & "SELECT LocationID, ClientID, LocationName AS Location, LAddrLine1, LState AS St, nz(FAUnit,0) + nz(Track,0) as Gar, Delv, Regd, Titled "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "AND  LocationID <> " & Me.LocID1 & " "
    SQLz = SQLz & "ORDER BY LAddrLine1, LState;"
    Me.LocationList2.RowSource = SQLz
    SQLz = ""
    SQLz = SQLz & "SELECT LAddrLine1 "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "AND  LocationID <> " & Me.LocID1 & " "
    SQLz = SQLz & "Group BY LAddrLine1 "
    SQLz = SQLz & "ORDER BY LAddrLine1 "
    Me.SearchAddress2.RowSource = SQLz
    ToggleSortName2 = 1
        Call SortListBox("ToggleSortName2", "LocationList2", "LocationName")
End Sub
Private Sub LocationList2_Click()
    Me.LocID2 = Me.LocationList2.Column(0)
End Sub

'Private Sub MoveUnits_Click()
'    Dim ItemIndex As Variant
'    Call DelWorkTable
'    For Each ItemIndex In Me.UnitsList.ItemsSelected
'        SQLz = "INSERT INTO MoveUnits_WorkTable ( UnitID ) SELECT " & Me.UnitsList.ItemData(ItemIndex) & " AS UnitID;"
'        DoCmd.SetWarnings False
'        DoCmd.RunSQL (SQLz)
'        DoCmd.SetWarnings True
'    Next
'    On Error GoTo fixme1
'    SQLz = "UPDATE Units INNER JOIN MoveUnits_WorkTable ON Units.UnitID = MoveUnits_WorkTable.UnitID SET Units.GroupREF = " & Me.NewGroupList & ";"
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQLz)
'    DoCmd.SetWarnings True
'    Me.UnitsFormsList.Requery
'    Me.UnitsList.Requery
'    Call CheckATFChanges(Me.NewGroupList, "Move Units")
'    Exit Sub
'fixme1:
'End Sub

Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From MoveUnits_WorkTable;")
    DoCmd.SetWarnings True
End Sub
Private Sub MoveGar_Click()
    If Me.LocID1 & 0 = 0 Then Exit Sub
    If Me.SelectClient & 0 = 0 Then Exit Sub
    addrx = LocationList1.Column(3) & "<br />, " & Me.LocationList1.Column(4)
    Typex = "Garaged"
    TargetX = "UnitGarageLocationRef"
    Call MoveUnits(5)
End Sub

Public Sub MoveDel_Click()
    If Me.LocID1 & 0 = 0 Then Exit Sub
    If Me.SelectClient & 0 = 0 Then Exit Sub
    addrx = LocationList1.Column(3) & "<br />, " & Me.LocationList1.Column(4)
    Typex = "Delivery"
    TargetX = "UnitDeliveryLocationRef"
    Call MoveUnits(6)
End Sub
Private Sub MoveReg_Click()
    If Me.LocID1 & 0 = 0 Then Exit Sub
    If Me.SelectClient & 0 = 0 Then Exit Sub
    addrx = LocationList1.Column(3) & "<br />, " & Me.LocationList1.Column(4)
    Typex = "Registration"
    TargetX = "UnitRegLocationRef"
    Call MoveUnits(7)
End Sub
Private Sub MoveTitle_Click()
    If Me.LocID1 & 0 = 0 Then Exit Sub
    If Me.SelectClient & 0 = 0 Then Exit Sub
    addrx = LocationList1.Column(3) & "<br />, " & Me.LocationList1.Column(4)
    Typex = "Titled"
    TargetX = "UnitTitleLocationRef"
    Call MoveUnits(8)
End Sub

Public Sub MoveUnits(Coli As Integer)
    msg1 = ""
    If Me.LocID2 & 0 = 0 Then
        Call Dialog.RichBox("You Must Select a<b> 'TO'</b> Address", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Exit Sub
    End If
    If Me.LocationList1.Column(Coli) < 1 Then
        Call Dialog.RichBox("You Must Select an Address That has <b>" & Typex & "</b> Units", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Exit Sub
    End If
    msg1 = msg1 & "<b><Font Size=3>Move Units From...</font></b><br/>"
    msg1 = msg1 & "Address:<br /><b>" & addrx & "</b><br/><br/>"
    msg1 = msg1 & "  <b>Are you sure?</b><br/><br/>"
    msg1 = msg1 & "Note:<br/>This will move " & Typex & " units from this address, only."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
    If ResposeX = vbBt2 Then Exit Sub
    SQLz = "Update Units Set " & TargetX & "=" & Me.LocID2 & " WHERE ClientID=" & Me.SelectClient & " AND " & TargetX & "=" & Me.LocID1 & ";"
    ' DoCmd.SetWarnings False
    On Error GoTo FixMe
    DoCmd.RunSQL (SQLz)
FixMe:
    On Error GoTo 0
    DoCmd.SetWarnings True
    Call RefreshTaxForms
End Sub

Public Sub SearchAddress1_Click()
    SQLz = ""
    SQLz = SQLz & "SELECT LocationID, ClientID, LocationName AS Location, LAddrLine1, LState AS St, Cint(nz(FAUnit,0)) + Cint(nz(Track,0)) as Gar, Delv, Regd, Titled "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "AND LAddrLine1 Like '*" & Me.SearchAddress1 & "*' "
    SQLz = SQLz & "ORDER BY LAddrLine1, LState;"
    Me.SearchAddress2 = ""
    Me.LocationList1.RowSource = SQLz
End Sub

Private Sub SearchAddress2_Click()
    SQLz = ""
    SQLz = SQLz & "SELECT LocationID, ClientID, LocationName AS Location, LAddrLine1, LState AS St, Cint(nz(FAUnit,0)) + Cint(nz(Track,0)) as Gar, Delv, Regd, Titled "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "AND LAddrLine1 Like '*" & Me.SearchAddress2 & "*' "
    SQLz = SQLz & "AND  LocationID <> " & Me.LocID1 & " "
    SQLz = SQLz & "ORDER BY LAddrLine1, LState;"
    Me.LocationList2.RowSource = SQLz
End Sub

Public Sub SelectClient_Click()
    Me.ClientGroupID = Me.SelectClient.Column(3)
    SQLz = ""
    SQLz = SQLz & "SELECT LocationID, ClientID, LocationName AS Location, LAddrLine1, LState AS St, Cint(nz(FAUnit,0)) + Cint(nz(Track,0)) as Gar, Delv, Regd, Titled "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "ORDER BY LAddrLine1, LState;"
    Me.LocationList1.RowSource = SQLz
    Me.LocationList2.RowSource = SQLz
    SQLz = ""
    SQLz = SQLz & "SELECT LAddrLine1 "
    SQLz = SQLz & "From vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.ClientGroupID & " "
    SQLz = SQLz & "Group BY LAddrLine1 "
    SQLz = SQLz & "ORDER BY LAddrLine1 "
    Me.SearchAddress1.RowSource = SQLz
    Me.SearchAddress2.RowSource = SQLz
    Call ClearSearches
    ToggleSortName1 = 1
    ToggleSortName2 = 1
    Call SortListBox("ToggleSortName1", "LocationList1", "LocationName")
    Call SortListBox("ToggleSortName2", "LocationList2", "LocationName")
End Sub
Private Sub ClearSearches()
    Me.LocID1 = ""
    Me.LocID2 = ""
    Me.SearchAddress1 = ""
    Me.SearchAddress2 = ""
End Sub

Private Sub ToggleSortName1_Click()
Call SortListBox("ToggleSortName1", "LocationList1", "LocationName")
End Sub
Private Sub ToggleSortAddress1_Click()
Call SortListBox("ToggleSortAddress1", "LocationList1", "LAddrLine1")
End Sub
Private Sub ToggleSortState1_Click()
Call SortListBox("ToggleSortState1", "LocationList1", "LState")
End Sub
Private Sub ToggleSortGar1_Click()
Call SortListBox("ToggleSortGar1", "LocationList1", "Cint(nz(FAUnit,0)) + Cint(nz(Track,0))")
End Sub
Private Sub ToggleSortDelv1_Click()
Call SortListBox("ToggleSortDelv1", "LocationList1", "Delv")
End Sub
Private Sub ToggleSortRegd1_Click()
Call SortListBox("ToggleSortRegd1", "LocationList1", "Regd")
End Sub
Private Sub ToggleSortTitled1_Click()
Call SortListBox("ToggleSortTitled1", "LocationList1", "Titled")
End Sub

Private Sub ToggleSortName2_Click()
Call SortListBox("ToggleSortName2", "LocationList2", "LocationName")
End Sub
Private Sub ToggleSortAddress2_Click()
Call SortListBox("ToggleSortAddress2", "LocationList2", "LAddrLine1")
End Sub
Private Sub ToggleSortState2_Click()
Call SortListBox("ToggleSortState2", "LocationList2", "LState")
End Sub
Private Sub ToggleSortGar2_Click()
Call SortListBox("ToggleSortGar2", "LocationList2", "Cint(nz(FAUnit,0)) + Cint(nz(Track,0))")
End Sub
Private Sub ToggleSortDelv2_Click()
Call SortListBox("ToggleSortDelv2", "LocationList2", "Delv")
End Sub
Private Sub ToggleSortRegd2_Click()
Call SortListBox("ToggleSortRegd2", "LocationList2", "Regd")
End Sub
Private Sub ToggleSortTitled2_Click()
Call SortListBox("ToggleSortTitled2", "LocationList2", "Titled")
End Sub

Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
