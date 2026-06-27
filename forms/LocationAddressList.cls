Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim RST        As DAO.Recordset
Dim SQLi, SQLx, SQLy, SQLz As String
Dim CGrpInt As Long

Private Sub ClearFilters_Click()
    Me.SearchState = ""
    Me.SearchAddress = ""
    Me.SearchCity = ""
    Me.SearchName = ""
    Me.LocationFilter1 = 0
    Me.LocationFilter2 = 0
    Me.SearchState.SetFocus
    Call SearchState_Click
End Sub

Public Sub SetForm()

    Me.LocationFilter1 = 0
    Me.LocationFilter2 = 0
    SQLz = ""
    SQLx = ""
    SQLz = SQLz & "Select * "
    SQLz = SQLz & "FROM vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID=" & CGrpInt & " "
    SQLy = SQLz
    SQLz = SQLz & "Order By Lstate, LCity"
    'Me.Form.RecordSource = SQLz
    SQLi = SQLi & "FROM vw_LocationAddressMatrix "
    SQLi = SQLi & "WHERE ClientGroupID=" & DLookup("ClientGroupID", "LocationAddresses", "LocationID=" & Me.Parent.LAddressID) & " "

    SQLx = "Select LCity "
    SQLx = SQLx & "FROM Locations "
    SQLx = SQLx & "WHERE ClientGroupID=" & Me.Parent.ClientGroupID & " "
    SQLx = SQLx & "GROUP BY LCity Order By LCity"
    Me.SearchCity.RowSource = SQLx

    SQLx = "Select LState "
    SQLx = SQLx & "FROM Locations "
    SQLx = SQLx & "WHERE ClientGroupID=" & Me.Parent.ClientGroupID & " "
    SQLx = SQLx & "GROUP BY LState Order By LState"
    Me.SearchState.RowSource = SQLx

    SQLx = "Select LAddrLine1 "
    SQLx = SQLx & "FROM Locations "
    SQLx = SQLx & "WHERE ClientGroupID=" & Me.Parent.ClientGroupID & " "
    SQLx = SQLx & "GROUP BY LAddrLine1 Order By LAddrLine1"
    Me.SearchAddress.RowSource = SQLx

    SQLx = "Select LocationName "
    SQLx = SQLx & "FROM Locations "
    SQLx = SQLx & "WHERE ClientGroupID=" & Me.Parent.ClientGroupID & " "
    SQLx = SQLx & "GROUP BY LocationName Order By LocationName"
    Me.SearchName.RowSource = SQLx
End Sub


Private Sub Form_Current()
    Me.txtCurrent = Me.LocationID
End Sub

Public Sub Form_Load()
    CGrpInt = Me.Parent.ClientGroupID
    Call SetForm
End Sub
Private Sub LAddrLine1_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub LCity_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub LocationName_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub LState_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub RentReceipt_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub SalesTaxRate_Click()
    Me.Parent.MoveTo = Me.LocationID
    Call Forms("LocationAddressesEdit").MoveRecordTo
End Sub

Private Sub SearchAddress_Click()
    Me.SearchState = ""
    Me.SearchCity = ""
    SQLx = SQLy
    SQLx = SQLx & "AND LAddrLine1 Like '*" & Me.SearchAddress & "*' "
    SQLx = SQLx & "Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
    Me.SearchAddress.SetFocus
End Sub
Private Sub SearchCity_Click()
    Me.SearchState = ""
    Me.SearchAddress = ""
    Me.SearchName = ""
    SQLx = SQLy
    SQLx = SQLx & "AND LCity Like '*" & SearchCity & "*'"
    SQLx = SQLx & "Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
End Sub

Private Sub SearchName_Click()
    Me.SearchAddress = ""
    Me.SearchCity = ""
    Me.SearchState = ""
    SQLx = SQLy
    SQLx = SQLx & "AND LocationName Like '*" & Me.SearchName & "*'"
    SQLx = SQLx & "Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
End Sub

Private Sub SearchState_Click()
    Me.SearchAddress = ""
    Me.SearchCity = ""
    Me.SearchName = ""
    SQLx = ""
    SQLx = SQLx & "Select * "
    SQLx = SQLx & "FROM vw_LocationAddressMatrix "
    SQLx = SQLx & "WHERE ClientGroupID=" & CGrpInt & " "
    SQLx = SQLx & "AND Lstate Like '*" & Me.SearchState & "*'"
    SQLx = SQLx & "Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
End Sub
Private Sub LocationFilter1_Click()
    Me.LocationFilter2 = 0
    SQLx = ""
    SQLx = SQLx & "Select * "
    SQLx = SQLx & "FROM vw_LocationAddressMatrix "
    SQLx = SQLx & "WHERE ClientGroupID=" & CGrpInt & " "
    Select Case LocationFilter1
        'FAUnit, Track, Regd, Delv, Titled"
        'Bill, Gar, HQ, Reg, Title "

    Case 1
        SQLx = SQLx & "AND FAUnit <> 0"
    Case 2
        SQLx = SQLx & "AND Track > 0"
    Case 3
        SQLx = SQLx & "AND Regd > 0"
    Case 4
        SQLx = SQLx & "AND Delv > 0"
    Case 5
        SQLx = SQLx & "AND Titled > 0"
    End Select
    If Me.SearchState <> "" Then SQLx = SQLx & " AND LState Like '*" & Me.SearchState & "*'"
    If Me.SearchAddress <> "" Then SQLx = SQLx & " AND LAddress1 Like '*" & Me.SearchAddress & "*'"
    If Me.SearchCity <> "" Then SQLx = SQLx & " AND LCity Like '*" & Me.SearchCity & "*'"
    If Me.SearchName <> "" Then SQLx = SQLx & "AND LocationName Like '*" & Me.SearchName & "*'"
    SQLx = SQLx & " Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
End Sub
Private Sub LocationFilter2_Click()
    Me.LocationFilter1 = 0
    SQLx = ""
    SQLx = SQLx & "Select * "
    SQLx = SQLx & "FROM vw_LocationAddressMatrix "
    SQLx = SQLx & "WHERE ClientGroupID=" & CGrpInt & " "
    Select Case LocationFilter2
        'FAUnit, Track, Regd, Delv, Titled"
        'Bill, Gar, HQ, Reg, Title "
    Case 6
        SQLx = SQLx & "AND Bill ='Y'"
    Case 7
        SQLx = SQLx & "AND Gar ='Y'"
    Case 8
        SQLx = SQLx & "AND HQ ='Y'"
    Case 9
        SQLx = SQLx & "AND Reg ='Y'"
    Case 10
        SQLx = SQLx & "AND Title ='Y'"
    End Select
    If Me.SearchState <> "" Then SQLx = SQLx & " AND LState Like '*" & Me.SearchState & "*'"
    If Me.SearchAddress <> "" Then SQLx = SQLx & " AND LAddress1 Like '*" & Me.SearchAddress & "*'"
    If Me.SearchCity <> "" Then SQLx = SQLx & " AND LCity Like '*" & Me.SearchCity & "*'"
    If Me.SearchName <> "" Then SQLx = SQLx & "AND LocationName Like '*" & Me.SearchName & "*'"
    SQLx = SQLx & " Order By Lstate, LCity"
    Me.Form.RecordSource = SQLx
End Sub
