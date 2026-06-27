Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub Form_Load()
SQLz = "SELECT Locations.LocationID, Locations.LocationName "
SQLz = SQLz & ",Locations.LAddrLine1, Locations.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Locations.LocObsdate, Locations.LocObsCmt"
SQLz = SQLz & " FROM Locations "
SQLz = SQLz & " WHERE (((Locations.LocObsolete) = No) And ((Locations.ClientGroupID) = " & CLng(Nz(Me.OpenArgs, 771)) & ")) "
SQLz = SQLz & " ORDER BY Locations.LocationName, Locations.LocationType;"
Me.XLocationID.RowSource = SQLz
End Sub

Private Sub SubmitMe_Click()
    Me.Visible = False
End Sub
