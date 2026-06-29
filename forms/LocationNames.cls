Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, SQLx As String
Dim RST        As DAO.Recordset
Private Sub Form_Load()
ToggleSortName = 1
ToggleSortState = 0
ToggleSortTax = 0
ToggleSortType = 0
ToggleSortUnits = 0
Me.ClientGroupID = Forms!ClientsView.ClientGroupID
End Sub
Private Sub Form_Current()
SQLz = "SELECT LAddressID, LAddrLine1 As Address, LState as State, LCity as City, LCounty as County, LPostalCode as Zip, LAddrName, ClientGroupID " & _
"FROM LocationAddresses " & _
"WHERE ClientGroupID =" & Me.ClientGroupID & " ORDER BY LState, LCity;"
Me.LAddressID.RowSource = SQLz
Call LocationID_Click
End Sub
Private Sub AddNew_Click()
If Me.ClientGroupID & 1 = 1 Then Me.ClientGroupID = Forms!ClientsView.ClientGroupID
DoCmd.close acForm, "LocationNamesAdd"
DoCmd.OpenForm "LocationNamesAdd", , , , , , CInt(Me.ClientGroupID)
Forms!LocationNamesAdd.ClientList = Me.ClientID
'On Error Resume Next
'If Me.Dirty Then Me.Dirty = False
'On Error GoTo FixMe
'    DoCmd.GotoRecord , , acNewRec
'Exit Sub
'FixMe:
'DoCmd.GotoRecord , , acNewRec
'On Error GoTo 0
End Sub

Private Sub ClearSorts_Click()
ToggleSortName = 1
ToggleSortState = 0
ToggleSortTax = 0
ToggleSortType = 0
ToggleSortUnits = 0
Me.LocationList.RowSource = SplitString(Me.LocationList.RowSource, "Order By", 0) & " Order By vw_LocationAddressMatrix.LocationName"
End Sub

Private Sub ContactsCBO_Click()
    Dim Msgx, ClassificationX As String
    Dim LinkTypeX As Integer
    Msgx = ""
    If IsNull(Me.LocationID) Then Msgx = Msgx & vbNewLine & "   Location"
    'If IsNull(Me.ContactsCBO.Column(0)) Then MsgX = MsgX & vbNewLine & "   Contact to Link"
    'If IsNull(Me.LinkTypeCBO.Column(0)) Then MsgX = MsgX & vbNewLine & "   Type of Link to use"
    If Msgx <> "" Then
        Msgx = "You must Select a" & Msgx
        MsgBox Msgx, vbCritical, "Missing Data"
        Exit Sub
    End If
    ClassificationX = DLookup("contactClassification", "Contacts", "contactID=" & Me.ContactsCBO.Column(0))
    LinkTypeX = DLookup("LinkTypeID", "ContactsLinkedToLocationTypes", "LinkType='" & ClassificationX & "'")
    If DLookup("ContactID", "ContactsLinkedToLocation", "ContactID=" & Me.ContactsCBO.Column(0) & " AND LinkTypeID =" & LinkTypeX & " AND LocationID=" & Me.LocationID) Then
        MsgBox Me.ContactsCBO.Column(2) & vbNewLine & "Is already linked  to " & Me.LocationName & vbNewLine & "as a '" & ClassificationX & "' contact type" & vbNewLine & vbNewLine & "Canceled...", vbCritical, "We've Done this already"

        Exit Sub
    End If
    SQLz = "INSERT INTO ContactsLinkedToLocation ( ContactID, LocationID, LinkTypeID ) " & _
           "SELECT " & Me.ContactsCBO.Column(0) & ", " & Me.LocationID & ", " & LinkTypeX & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedContactsList.Requery
    Me.Parent.ClientContact_sub.Form.LinkedLocationsList.Requery
    Call SetUnits
    Me.ContactDetail = ""
End Sub

Private Sub LinkedContactsList_Click()
'    SQLz = "Select * From Contacts WHERE ContactID =" & Me.LinkEdContactsList.Column(0) & ";"
'    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
'    Me.ContactDetail = RST![contactFullName] & vbNewLine & RST!contactTitle & vbNewLine & _
'                       LAddrLine1 & vbNewLine & _
'                       Me.[LCity] & " " & Me.LState & " " & Me.LZip & vbNewLine & _
'                       "Phone: " & Left(RST![contactPrimaryPhone], 3) & "-" & Mid(RST![contactPrimaryPhone], 4, 3) & "-" & Right(RST![contactPrimaryPhone], 4) & vbNewLine & _
'                       "Email: " & RST![contactEmail] & vbNewLine
'    RST.close
'    Set RST = Nothing
End Sub

Private Sub LocationsReportPrintView_Click()
DoCmd.close acReport, "LocationDirectory"
DoCmd.OpenReport "LocationDirectory", acViewPreview
End Sub

Private Sub LocationsReportView_Click()
DoCmd.close acReport, "LocationDirectory"
DoCmd.OpenReport "LocationDirectory", acViewReport
End Sub

Private Sub LookUpTaxZip_Click()
    Dim Zipx   As String
    Dim StateX As String
    StateX = Nz(DLookup("LState", "LocationAddresses", "LAddressID=" & LAddressID), "")
    Zipx = Nz(DLookup("LPostalCode", "LocationAddresses", "LAddressID=" & LAddressID), "")
    If StateX & "test" = "test" Or Zipx & "test" = "test" Then
        MsgBox "An Address must have both a State and a Zip," & vbNewLine & "in order to perform a 'Look Up Tax (Zip)'", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Call LookUpTaxByZipWeb(StateX, Zipx)
End Sub

Private Sub SaveMe_Click()
If Me.Dirty Then Me.Dirty = False
Me.LocationList.Requery
'Me.Form.Requery
Me.LocationList = Me.LocationID

End Sub

Private Sub EditState_Click()
DoCmd.close acForm, "LocationAddressesEdit"
If Me.LAddressID & "test" = "test" Then
DoCmd.OpenForm "LocationAddressesEdit", acNormal, , "ClientGroupID=" & Me.ClientGroupID
Else
DoCmd.OpenForm "LocationAddressesEdit", acNormal, , "LAddressID=" & Me.LAddressID
End If
End Sub

Private Sub LocationID_Click()
Dim TempString As String
   On Error GoTo FixMe
SQLz = "Select * From LocationAddresses WHERE LAddressID=" & Me.LAddressID.Column(0) & ";"
Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
If Not IsNull(RST!LAddrLine2) Then TempString = TempString & RST!LAddrLine2 & "<br />"
TempString = TempString & RST!LCity & ", " & RST!LState & " " & RST!LPostalCode & "<br />"
TempString = TempString & "<b>County</b> " & RST!LCounty
Me.AddressText = TempString

Set RST = Nothing
FixMe:
End Sub

Public Sub LocationList_Click()
    On Error GoTo FixMe
    Dim LocX   As Integer
    LocX = Me.LocationList.Column(0)
    Set RST = Me.RecordsetClone
    RST.FindFirst "LAddressID = " & LocX
    Me.Bookmark = RST.Bookmark
    Call SetUnits
    RST.close
    Set RST = Nothing
    Me.ContactDetail = ""
  '  Me.LocationList = LocX

FixMe:
End Sub
Private Sub SetUnits()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""

    SQLz = "SELECT AssetID, iif(vw_SixKeys.MLOrig='FA','FA','Track')As MLOrig, iif(vw_SixKeys.MLOrig='FA',Schedule & ' - ' & UnitGroup, Schedule)as Sch_Grp, Unitnum, UnitDescShortVer AS [Desc], Status AS Sts "
    SQLz = SQLz & "FROM vw_SixKeys WHERE vw_SixKeys.LocationID=" & Me.LocationList.Column(0) & " "
    SQLz = SQLz & "ORDER BY MLOrig, SchID, sGrpID, Status, UnitDescShortVer;"
    'SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc, UnitStatus as STS From " & SourceUnitQueryX & " " & _
           "Where GarageLocationID=" & Me.LocationID & " Order by MLOrig, AssetID;"
    Me.UnitsList.RowSource = SQLz

    Dim ListCountX As String
    ListCountX = "0 Units @ Location"
    If Me.UnitsList.ListCount > 0 Then ListCountX = Me.UnitsList.ListCount - 1 & " Units @ Location"
    If (Me.UnitsList.ListCount - 1) = 1 Then ListCountX = Me.UnitsList.ListCount - 1 & " Unit @ Location"
    Me.UNITS.caption = ListCountX

    ListCountX = "0 Contacts Linked to Location"
    If Me.LinkedContactsList.ListCount > 0 Then ListCountX = Me.LinkedContactsList.ListCount - 1 & " Contacts Linked to Location"
    If (Me.LinkedContactsList.ListCount - 1) = 1 Then ListCountX = Me.LinkedContactsList.ListCount - 1 & " Contact Linked to Location"
    Me.ContactsLinked.caption = ListCountX
End Sub


Private Sub ToggleSortState_Click()
Call SortListBox("ToggleSortState", "LocationList", "vw_LocationAddressMatrix.LState")
End Sub

Private Sub ToggleSortName_Click()
Call SortListBox("ToggleSortName", "LocationList", "vw_LocationAddressMatrix.LocationName")
End Sub

Private Sub ToggleSortTax_Click()
Call SortListBox("ToggleSortTax", "LocationList", "Format([SalesTaxRate],'#.000%')")
End Sub

Private Sub ToggleSortType_Click()
Call SortListBox("ToggleSortType", "LocationList", "IIf([bill]='Y','-Bill',Null) & IIf([Gar]='Y','-Gar',Null) & IIf([HQ]='Y','-HQ',Null) & IIf([Reg]='Y','-Reg',Null) & IIf(vw_LocationAddressMatrix.[Title]='Y','-Til',Null)")
End Sub

Private Sub ToggleSortUnits_Click()
Call SortListBox("ToggleSortUnits", "LocationList", "IIf([FAUnit]>0,' FA:' & Nz([FAUnit]),Null) & IIf([Track]>0,' Tk:' & Nz([Track]),Null)")
End Sub
Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then Me(Ccrl.Name) = 0
        '    If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
        '    If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        'End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function

Private Sub UnitsList_DblClick(Cancel As Integer)
    If Me.UnitsList.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitsList.Column(0) & ""
    Call OpenUnitsForm(Me.UnitsList.Column(0))
    Me.LocationList.SetFocus
End Sub

Private Sub UnLinkedContact_Click()
    If IsNull(Me.LinkedContactsList.Column(0)) Then Exit Sub
    Dim Askit  As String
    Askit = MsgBox("You're about to Un-Link..." & vbNewLine & "Name: " & Me.LinkedContactsList.Column(3) & vbNewLine & "From:  " & Me.LocationName & vbNewLine & "Type:   " & Me.LinkedContactsList.Column(2) & vbNewLine & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Un-Link")
    If Askit <> 6 Then Exit Sub
    SQLz = "Delete From ContactsLinkedToLocation " & _
           "WHERE  ContactID=" & Me.LinkedContactsList.Column(0) & " AND LinkTypeID=" & Me.LinkedContactsList.Column(5) & " AND LocationID=" & Me.LocationID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedContactsList.Requery
    Me.Parent.ClientContact_sub.Form.LinkedLocationsList.Requery
    Call SetUnits
    Me.ContactDetail = ""
End Sub
