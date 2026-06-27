Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST As DAO.Recordset
Dim SQLz As String


Private Sub AddNew_Click()
DoCmd.GotoRecord , , acNewRec
End Sub

Private Sub Form_Current()
Me.LocationList.Requery
Me.LocationList = Me.LocationID
Call LocationList_Click
    Me.SearchAssetId = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
SQLz = "SELECT DivisionID, DivisionName, ClientID From ClientDivisions " & _
       "WHERE Deleted=False and ClientID=" & Me.clientID & " Order by DivisionName"
Me.DivisionID.RowSource = SQLz
End Sub

Private Sub LocationList_Click()
On Error GoTo FixMe
'Me.Form.FilterOn = False
'Me.Filter = "LocationID = " & Me.LocationList.Column(0)
'Me.Form.FilterOn = True
    Set RST = Me.RecordsetClone
    RST.FindFirst "LocationID = " & Me.LocationList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    RST.close
    Set RST = Nothing
SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " Order by MLOrig, AssetID;"
Me.UnitsList.RowSource = SQLz
FixMe:
End Sub



Private Sub SearchAssetId_Change()
    'Me.SearchAssetId = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND AssetID Like '*" & Me.SearchAssetId.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchLease_Change()
    Me.SearchAssetId = ""
    Me.SearchUnit = ""
    'Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND MLNo Like '*" & Me.SearchLease.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchSchedule_Change()
    Me.SearchAssetId = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    'Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND Schedule Like '*" & Me.SearchSchedule.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchUnit_Change()
    Me.SearchAssetId = ""
    'Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND [Unit#] Like '*" & Me.SearchUnit.Text & "*' Order by [Unit#];"
    Me.UnitsList.RowSource = SQLz
End Sub

Private Sub UnitsList_DblClick(Cancel As Integer)
If Me.UnitsList.Column(0) & "test" = "test" Then
MsgBox "There must be a unit shown first before you can select it", , "No Results"
Exit Sub
End If
    Call OpenUnitsForm(Me.UnitsList.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitsList.Column(0) & ""
    Me.LocationList.SetFocus
End Sub
