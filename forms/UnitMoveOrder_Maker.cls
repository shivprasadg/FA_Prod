Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, Commit   As String
Public ReportX, Queryx As String
Public ThePersonX, TheEmailX, TheLocationX As String
Public ThePhoneNumX As Variant

Private Sub CloseMe_Click()
    Call AutoUpdateEOL_Statuses(True)
    'Call FixParkedUnits(True)
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Close()
    Call AutoUpdateEOL_Statuses(True)
    DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub Form_Load()
    Me.Search_Client = ""
    Me.Search_Client.SetFocus
    Me.LocationList.RowSource = ""
    Me.UnitsList.Visible = False
    Me.UnitListcount.Visible = False
    Me.line2.Visible = False
    Me.MoveByUnit.Visible = False
    'Me.SearchVin.Visible = False
    SQLz = ""
    SQLz = SQLz & "SELECT ClientID, ClientGroupID,ClientShortName As Client, Count(AssetID) AS Units "
    SQLz = SQLz & "From vw_SixKeys "
    SQLz = SQLz & " WHERE (Status='EG' Or Status='EM')"
    SQLz = SQLz & " OR (Status='EZ' AND ParkRepairID Is Null) "
    SQLz = SQLz & " OR (Status='R' AND ParkRepairID Is Null)  "
    SQLz = SQLz & " GROUP BY ClientShortName, ClientGroupID, ClientID;"

    Me.ClientList.RowSource = SQLz
    ReportX = "MoveOrderPickUpPage"
    Queryx = "vw_EOLDashboard"
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub
Private Sub Location_AllOrdred_Click()
    If Me.LocationList.Column(0) & "test" = "test" Then Exit Sub
    Committ = MsgBox("You are about to mark all 'Not Ordered' as 'ORDERED'" & vbNewLine & "Do you Agree?", vbQuestion + vbYesNo + vbCritical, "All Ordered?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update UnitMoveOrder "
    SQLz = SQLz & "Set mStatus= 'Ordered' "
    SQLz = SQLz & "WHERE mStatus='Not Ordered' and mOldLocationID = " & Me.LocationList.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.UnitsList.Requery
    Me.LocationList.Requery
    Me.Form.Recalc
End Sub

Public Sub ClientList_Click()
    Me.LocationList.SetFocus
    Me.LocationList = ""
    Me.UnitsList.Visible = False
    Me.UnitListcount.Visible = False
    Me.line2.Visible = False
    'Me.SearchVin.Visible = False
    'Me.SearchVin = ""
    Me.MoveByUnit.Visible = False
    Me.ClientGroupID = Me.ClientList.Column(1)
    SQLz = ""
    SQLz = SQLz & "SELECT mOldLocationID as LocationID, PickupLocation as Pickup, mStatus, Count(AssetID) as Units "
    SQLz = SQLz & "FROM vw_UnitMoveDashboard AS M "
    SQLz = SQLz & "WHERE (Status ='EG' AND mClientID=" & Me.ClientList.Column(0) & ") "
    SQLz = SQLz & "OR (Status='EM' AND mClientID=" & Me.ClientList.Column(0) & ") "
    SQLz = SQLz & "OR (Status='EZ' AND DeliveryLocation Is Null AND mClientID=" & Me.ClientList.Column(0) & ") "
    SQLz = SQLz & "OR (Status='R' AND DeliveryLocation Is Null AND mClientID=" & Me.ClientList.Column(0) & ") "
    SQLz = SQLz & "Group BY mOldLocationID, PickupLocation, mStatus "
    SQLz = SQLz & "ORDER BY PickupLocation;"
    Me.LocationList.RowSource = SQLz
    Call ClearList(Me.LocationList.Name)
    Call ClearForm
    Call ClearToggles(0)
End Sub

Public Sub LocationList_Click()
    Me.UnitListcount.Visible = True
    Me.line2.Visible = True
    Me.UnitsList.Visible = True
    'Me.SearchVin.Visible = True
    Me.MoveByUnit.Visible = True
    Me.UnitsList = ""
    'Me.SearchVin = ""
    ReportX = "MoveOrderRelease"
    SQLz = ""
    SQLz = SQLz & "SELECT AssetID, UnitNum, VIN, [Status],[MStatus], UnitDescShortVer as uDesc, mStartDate as Start, DeliveryLocation as [Deliver To] "
    SQLz = SQLz & "FROM vw_UnitMoveDashboard "
    SQLz = SQLz & "WHERE (Status='EG' AND mOldLocationID=" & [LocationList] & ")  "
    SQLz = SQLz & "OR (Status='EM' AND mOldLocationID=" & [LocationList] & " ) "
    SQLz = SQLz & "OR (Status='EZ' AND DeliveryLocation Is Null AND mOldLocationID=" & [LocationList] & " ) "
    SQLz = SQLz & "OR (Status='R' AND DeliveryLocation Is Null AND mOldLocationID=" & [LocationList] & " ) "
    SQLz = SQLz & "ORDER BY mStartDate, UnitNum,VIN;"
    Me.UnitsList.RowSource = SQLz
    Call ClearForm
    Call ClearToggles(0)
End Sub

Private Sub SubmitMoveOrder_Click()
    If Me.mNewLocationID & "test" = "test" Or Me.mAssetId & "test" = "test" Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Update Units Set UnitStatus='EM', LocationParkedID= " & Me.mNewLocationID & " WHERE UnitID=" & Me.mAssetId & ";")
    DoCmd.SetWarnings True
End Sub

Private Sub ToggleSortmStatus_Click()
   Call SortListBox("ToggleSortmStatus", "LocationList", "mStatus")
End Sub

Private Sub ToggleSortPickup_Click()
   Call SortListBox("ToggleSortPickup", "LocationList", "PickupLocation")
End Sub
Private Sub ToggleSortUnits_Click()
   Call SortListBox("ToggleSortUnits", "LocationList", "Count(AssetID)")
End Sub
Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*Sort*" And Me(Ccrl.Name).Name Like "*Sort*" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Public Sub ClearToggles(WhichToggles As Integer)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" Then
            If WhichToggles = 1 And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If WhichToggles = 2 And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
            If WhichToggles = 0 Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
End Sub


Public Sub UnitsList_Click()
    If DCount("mAssetID", "UnitMoveOrder", "mAssetID=" & Me.UnitsList.Column(0)) = 1 Then
        SQLz = "Select * From UnitMoveOrder Where mAssetID =" & Me.UnitsList.Column(0)
        Me.Form.RecordSource = SQLz
    Else
        SQLz = "Select * From UnitMoveOrder Where mAssetID =0"
        Me.Form.RecordSource = SQLz

        Committ = MsgBox("Do you want to add a Driveaway Move for this Unit: " & vbNewLine & Me.UnitsList.Column(5) & vbNewLine & "VIN: " & Me.UnitsList.Column(2) & "?", vbQuestion + vbYesNo + vbCritical, "Move Unit?")
        If Committ = vbNo Then Exit Sub
        If Nz(DLookup("LocationParkedID", "Units", "UnitID=" & Me.UnitsList.Column(0) & ""), 0) = 0 Then GoTo NOMatrix

        Dim RST        As DAO.Recordset
        Set RST = CurrentDb.OpenRecordset("SELECT * FROM UnitMoveOrderMatrix WHERE ClientGroupId=" & Me.ClientGroupID & " AND ParkedLocID =" & DLookup("LocationParkedID", "Units", "UnitID=" & Me.UnitsList.Column(0) & "") & ";", dbOpenSnapshot)
        If RST.RecordCount > 0 Then
            Call ParkedLocationMove(RST!ParkedLocID, RST!RMKTDepotID, Me.UnitsList.Column(0))
        Else
NOMatrix:
            If Nz(DLookup("LocationParkedID", "Units", "UnitID=" & Me.UnitsList.Column(0) & ""), 0) <> 0 Then
                Call ParkedLocationMove(DLookup("LocationParkedID", "Units", "UnitID=" & Me.UnitsList.Column(0) & ""), DLookup("LocationParkedID", "Units", "UnitID=" & Me.UnitsList.Column(0) & ""), Me.UnitsList.Column(0))
            Else
                Call ParkedLocationMove(DLookup("UnitGarageLocationRef", "Units", "UnitID=" & Me.UnitsList.Column(0) & ""), getValueFromPopUp("PopUp_PickLocation", "XLocationID", "771"), Me.UnitsList.Column(0), , , , True)
            End If
        End If
        ' DoCmd.GotoRecord , , acNewRec
        SQLz = "Select * From UnitMoveOrder Where mAssetID =" & Me.UnitsList.Column(0)
        Me.Form.RecordSource = SQLz
        Me.UnitsList.Requery
        Me.LocationList.Requery
    End If
        Me.mOldLocationID.Requery
        Me.mNewLocationID.Requery
End Sub
Private Sub Search_Client_Change()
    SQLz = "SELECT ClientID"
    SQLz = SQLz & ",ClientGroupID"
    SQLz = SQLz & ",ClientShortName AS Client"
    SQLz = SQLz & ",Count(vw_SixKeys.AssetID) As Units"
    SQLz = SQLz & " FROM  vw_SixKeys "
    SQLz = SQLz & " WHERE (((Status) = 'EG'"
    SQLz = SQLz & " OR (Status) = 'EM'))"
    SQLz = SQLz & " OR (((Status) = 'EZ') AND ((ParkRepairID) IS NULL))"
    SQLz = SQLz & " GROUP BY ClientID, ClientGroupID, ClientShortName "
    SQLz = SQLz & " HAVING ClientShortName Like '*" & Me.Search_Client.Text & "*'"
    SQLz = SQLz & " ORDER BY ClientShortName"
    Me.ClientList.RowSource = SQLz
End Sub

Private Sub SearchVIN_Change()
    SQLz = ""
    SQLz = SQLz & "SELECT AssetID, UnitNum, VIN, [Status],[MStatus], UnitDescShortVer as uDesc, mStartDate as Start, DeliveryLocation as [Deliver To] "
    SQLz = SQLz & "FROM vw_UnitMoveDashboard "
    SQLz = SQLz & "WHERE (Status='EG' AND mOldLocationID=" & [LocationList] & " AND VIN Like '*" & Me.SearchVin.Text & "*' )"
    SQLz = SQLz & "OR (Status='EM' AND mOldLocationID=" & [LocationList] & " AND VIN Like '*" & Me.SearchVin.Text & "*' )"
    SQLz = SQLz & "OR (Status='EZ' AND DeliveryLocation Is Null AND mOldLocationID=" & [LocationList] & " AND (VIN Like '*" & Me.SearchVin.Text & "*' )"
    SQLz = SQLz & "OR (Status='R' AND DeliveryLocation Is Null AND mOldLocationID=" & [LocationList] & " AND (VIN Like '*" & Me.SearchVin.Text & "*' )"
    SQLz = SQLz & " ORDER BY mStartDate, UnitNum,VIN;"
    Me.UnitsList.RowSource = SQLz
End Sub

Private Sub MoveByUnit_Click()
    If Me.UnitsList & "test" = "test" Then
        MsgBox "Please Select a Unit", vbInformation, "Nothing = Nothing"
        Exit Sub
    End If
    Call MakePickUp("Unit")
End Sub
Private Sub PickUpByLocation_Click()
    If Me.LocationList & "test" = "test" Then
        MsgBox "Please Select a Location", vbInformation, "Location, Location, Location"
        Exit Sub
    End If
    Call MakePickUp("Location")
End Sub



Public Sub MakePickUp(ReportType As String)
    Call CloseReports
    SQLz = "SELECT * FROM vw_UnitMoveDashboard "
    ReportX = "MoveOrderPickUpPage"
    If ReportType = "Unit" Then SQLz = SQLz & " WHERE AssetID = " & Me.UnitsList.Column(0) & " "
    If ReportType = "Location" Then
        If DCount("AssetID", "vw_UnitMoveDashboard", "mStatus ='Not Ordered'AND mOldLocationID = " & Me.LocationList.Column(0) & " ") = 0 Then
            Committ = MsgBox("There are no 'To Move' Units at this location: " & vbNewLine & "do you still want to see the Move Order?", vbQuestion + vbYesNo + vbCritical, "Still see it?")
            If Committ = vbNo Then Exit Sub
            SQLz = SQLz & " WHERE mStatus ='Ordered' AND mOldLocationID = " & Me.LocationList.Column(0) & " "
        Else
            SQLz = SQLz & " WHERE mStatus ='Not Ordered' AND mOldLocationID = " & Me.LocationList.Column(0) & " "
        End If
    End If
    SQLz = SQLz & " Order By AssetID;"
    DoCmd.OpenReport ReportX, acViewPreview, , , , SQLz
End Sub
Private Sub CloseReports()

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
Public Sub ClearForm()
    SQLz = "Select * From UnitMoveOrder Where mAssetID =0"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub UnitsList_DblClick(Cancel As Integer)
    If Me.UnitsList.Column(0) & "test" = "test" Then
        MsgBox "There must select a unit before you can view it", , "Really?"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitsList.Column(0))
    DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitsList.Column(0) & ""
    Me.Search_Client.SetFocus
End Sub
