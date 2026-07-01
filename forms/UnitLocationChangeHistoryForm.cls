Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim HS, SQLz As String
Public Function LoadData(UnitIDx As Long)
On Error Resume Next
    If Nz(Me.AssetID, 0) = 0 Or Nz(Me.AssetID, 0) <> UnitIDx Then
    SQLz = " SELECT H.*, LOld.LocationName AS PrevLocation, LNew.LocationName AS NewLocation "
    SQLz = SQLz & "FROM (LocationChangeHistory AS H "
    SQLz = SQLz & "INNER JOIN Locations AS LOld ON H.OldLocationId = LOld.LocationID) "
    SQLz = SQLz & "INNER JOIN Locations AS LNew ON H.NewLocationId = LNew.LocationID "
    SQLz = SQLz & "WHERE AssetID =" & UnitIDx & ";"
    Me.Form.RecordSource = SQLz
    End If
    On Error GoTo 0
    Call SetLocks
End Function
Private Sub SetLocks()
    HS = "Hide"
    Call ShrinkGrow
    If FAUserAccess > 24 Then
        Me.LocationChangeDate.Locked = False
        Me.Reason.Locked = False
    Else
        Me.LocationChangeDate.Locked = True
        Me.Reason.Locked = True
    End If
End Sub

Private Sub DeleteRecord_Click()
On Error GoTo FixMe
  DoCmd.RunSQL ("Delete From LocationChangeHistory WHERE ID = " & Me.ID & ";")
  Me.Form.Refresh
FixMe:
End Sub
Private Sub Form_Load()
  If FAUserAccess > 69 Then Me.DeleteRecord.Visible = True Else Me.DeleteRecord.Visible = False
End Sub
Private Sub HideShow_Click()
Call ShrinkGrow
End Sub

Private Sub HideShow_GotFocus()
If HS = "Show" Then HS = "Hide" Else HS = "Show"
Call ShrinkGrow
End Sub


Private Sub ShrinkGrow()
If HS = "Show" Then
HS = "Hide"
Me.Detail.Height = 500
Me.Reason.Height = 260
Me.FaUserName.Height = 260
Else
Me.Reason.Height = 0
Me.FaUserName.Height = 0
Me.Detail.Height = 250
HS = "Show"
End If
End Sub
