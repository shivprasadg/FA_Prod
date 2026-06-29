Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLz, SourceLocQueryX, SourceUnitQueryX As String


Private Sub LCity_Click()
    Dim CityReply As String
    Dim NumRecords As Integer
    If Me.LCity Like "*Unkn*" Then
        Me.LCity = "Unknown"
        Me.LCityID = 0
        GoTo SaveMe
    End If
    CityReply = LookUpByCity(Me.LCity, Me.LState)
    If CityReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(CityReply, "|", 0))
    If Me.LCounty <> CheckCounty(Nz(Me.LCounty, 0), Me.LState, Me.LCity) Then
        If Me.LCounty <> "Unknown" Then
            MsgBox "The city '" & Me.LCity & "'" & vbNewLine & "is not normaly listed to be within " & Me.LCounty & " county" & vbNewLine & "(the County was corrected accordingly)", vbInformation, AtlaasVersion
            Me.LCounty.SetFocus
        End If
    End If
    If NumRecords = 0 Then
        MsgBox "You've entered a City that is not in the ATLAAS Core, Would you like to add it?", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Me.LCityID = Me.LCity.Column(1)
    Me.LCounty = Me.LCity.Column(2)
    Me.LCountyID = Me.LCity.Column(3)
    Me.LCity.RowSource = SplitString(CityReply, "|", 4)
    Me.LCounty.RowSource = SplitString(CityReply, "|", 5)
SaveMe:
    If Me.Form.Dirty Then Me.Form.Dirty = False
    'Call LCounty_Click
End Sub

Private Sub LState_Click()
    If Nz(Me.LState, "XX") = "XX" Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If

   ' If Me.LState.Column(1) Like "*RR*" Then Me.LTaxOption = "RR"
   ' Me.LocationTax.Requery
    Dim StateReply As String
    Dim NumRecords As Integer
    Dim Msgx, ResposeX As String
    StateReply = Nz(LookUpByState(Me.LState), "0|")
    NumRecords = CInt(SplitString(StateReply, "|", 0))
    Me.LStateID = Me.LState.Column(3)
    Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
    Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
    If DCount("CityID", "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND City='" & Me.LCity & "'") = 0 Then Msgx = Msgx & "<br />The City '<b>" & Me.LCity & "</b>' is not within the State of <b> " & Me.LState & "</b>"
    If DCount(Nz(Me.LCountyID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CountyID =" & Nz(Me.LCountyID, 0)) = 0 Then Msgx = Msgx & "<br />The County '<b>" & Me.LCounty & "</b>' is not within the Sate of <b> " & Me.LState & "</b>"

    'If Me.LCity <> "Unknown" And DCount(Nz(Me.LCityID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CityID =" & Nz(Me.LCityID, 0)) = 0 Then Msgx = Msgx & "<br />The City '<b>" & Me.LCity & "</b>' is not within the State of <b> " & Me.LState & "</b>"
    'If Me.LCounty <> "Unknown" And DCount(Nz(Me.LCountyID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CountyID =" & Nz(Me.LCountyID, 0)) = 0 Then Msgx = Msgx & "<br />The County '<b>" & Me.LCounty & "</b>' is not within the Sate of <b> " & Me.LState & "</b>"
    If Msgx <> "" And Nz(Me.LCity, "") <> "" Then
        Msgx = "There is a address conflict:" & Msgx
        Msgx = Msgx & "<br /><br />If you press 'Proceed' ATLAAS will mark the unmatched fields to <b>'Unknown'</b><br />You than can select the proper values, afterwards"
        ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Undo", "Leave it", "Proceed")
        If ResposeX = 101 Then Me.Undo
        ' If ResposeX = vbBt2 Then
        If ResposeX = 103 Then
            If Msgx Like "*The City*" Then Me.LCity = "Unknown"
            If Msgx Like "*The City*" Then Me.LCityID = 0
            If Msgx Like "*County*" Then Me.LCounty = "Unknown"
            If Msgx Like "*County*" Then Me.LCountyID = 0
        End If
        If Me.Form.Dirty Then Me.Form.Dirty = False
    End If
    ' Me.LCounty = CheckCounty(Me.LCounty, Me.LState, Me.LCity)
    ' Me.LCity = CheckCity(Me.LCounty, Me.LState, Me.LCity)
    If NumRecords = 0 Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    If Me.Form.Dirty Then Me.Form.Dirty = False
End Sub


Private Sub SaveMe_Click()
    If Me.Dirty Then Me.Dirty = False
    Me.Form.Requery
    'Me.LocationList.Requery
End Sub
Private Sub ToggleSortName_Click()
    Call SortListBox("ToggleSortName", "LocationList", "LocationName")
End Sub
Private Sub ToggleSortState_Click()
    Call SortListBox("ToggleSortState", "LocationList", "LState")
End Sub
Private Sub ToggleSortTax_Click()
    Call SortListBox("ToggleSortTax", "LocationList", "LTaxRate")
End Sub
Private Sub ToggleSortType_Click()
    Call SortListBox("ToggleSortType", "LocationList", "LocationType")
End Sub
Private Sub ToggleSortUnits_Click()
    Call SortListBox("ToggleSortUnits", "LocationList", "Count(UnitId)")
End Sub
Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then Me(Ccrl.Name) = 0
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Private Sub ClearSorts_Click()
    ToggleSortName = 1
    ToggleSortState = 0
    ToggleSortTax = 0
    ToggleSortType = 0
    ToggleSortUnits = 0
    Me.LocationList.RowSource = SplitString(Me.LocationList.RowSource, "Order By", 0) & " Order By LocationName"
End Sub

Private Sub AddNew_Click()
'On Error Resume Next
    If Me.Dirty Then Me.Dirty = False
    On Error GoTo FixMe
    DoCmd.GotoRecord , , acNewRec
    Exit Sub
FixMe:
    DoCmd.GotoRecord , , acNewRec
    On Error GoTo 0
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Locations_sub"
End Sub

Public Sub SetLocations()
    Me.LocationList.RowSource = ""
    Me.LocationList.Requery
    On Error GoTo Form_Current_Error

    Dim ClientX, ClientGroupx As Long
    ClientX = Nz(Me.Parent.BankID, 0)
    'ClientGroupx = Nz(Me.Parent.ClientGroupID, 0)

    SQLz = "SELECT L.LocationID, L.LocationType AS Type, L.LocationName AS Name, L.LState AS State, Count(U.UnitID) AS Units"
    SQLz = SQLz & ", Format([LTaxRate],'Percent') AS TaxRate, L.LocObsolete, L.ClientREF "
    SQLz = SQLz & "FROM Locations AS L LEFT JOIN Units AS U ON L.LocationID = U.UnitGarageLocationRef "
    SQLz = SQLz & "WHERE L.ClientRef =" & ClientX & " "
    SQLz = SQLz & "GROUP BY L.LocationID, L.LocationType, L.LocationName, L.LState, L.LTaxRate, L.LocObsolete, L.ClientREF "
    SQLz = SQLz & "ORDER BY L.LocationType, L.LocationName, L.LocObsolete;"

    Me.LocationList.RowSource = SQLz
    On Error Resume Next
    Me.LocationList.Selected(0) = True
    On Error GoTo Form_Current_Error
    Call LocationList_Click
    'Me.ContactDetail = ""
    Exit Sub
Form_Current_Error:
    Dim m      As String
    If DCount("LocationID", "Locations", "ClientRef=" & Me.Parent.BankID & "") = 0 Then
        SQLz = "Insert into Locations (ClientRef, LocationType, LocationName, Locationsn, LocationStreet) Select " & Me.Parent.BankID & ",'Billing','Billing','Billing','" & Me.Parent.BankName & "';"
        m = "Error! But Don't Panic, Atlaas Knows what to do..." & vbNewLine
        m = m & "There isn't a 'Billing/Remit Address' assigend to Bank:" & Me.Parent.BankName & vbNewLine
        m = m & "Atlaas Core will assign a Remit Address for you." & vbNewLine & vbNewLine & "Please Press 'OK'"
        MsgBox m, vbInformation, "Atlaas Truly Loves You"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Call SetLocations
    End If
    Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Current of VBA Docume"
End Sub


Public Sub LocationList_Click()
'On Error GoTo Fixme
    Dim LocX   As Integer
    LocX = Me.LocationList.Column(0)
    SQLz = "SELECT Locations.*, lookup_States.StTaxType " & _
           "FROM Locations LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev " & _
           "WHERE Locations.LocationID = " & LocX & ";"
    Me.Form.RecordSource = SQLz

    Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
    Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
    Me.LocationList = LocX
FixMe:
End Sub
