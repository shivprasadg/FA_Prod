Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx, SQLz As String
Dim NeedSearch As Integer

Private Sub ClearFilters_Click()
    Call Form_Load
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DoSearch_Click()
    Call SetSQLx
    Call SetWhere
    NeedSearch = 0
    Me.ReLeaseUnitsSub.Form.RecordSource = SQLx & SQLz
End Sub

Private Sub ExportToExcel_Click()
    If NeedSearch = 1 Then Call DoSearch_Click
    Dim SQLv As String
    Dim XLS, TitleX As String
    Call SetSQLx
    Call SetWhere
    SQLv = SQLx & SQLz
'    XLS = "|A,TL|B,BL|E,$E|H,TX|J,BC|K,TC|M,BC|"
    MessageUser "Open", "Exporting ReLeased Units", "...Getting Data..."
    SendReport2Excel SQLv, "ReLeased Units", "ReLeased Units Export: " & Date, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub Form_Load()
    Me.AssetSearch = ""
    Me.BankSearch = ""
    Me.ClientSearch = ""
    Me.TitleNumSearch = ""
    Me.StartDate = ""
    Me.EndDate = ""
    Me.VINSearch = ""


    Call SetSQLx
    Call SetWhere
    Me.ReLeaseUnitsSub.Form.RecordSource = SQLx & SQLz
End Sub

Public Sub SetSQLx()
    SQLx = "SELECT * "
    SQLx = SQLx & "FROM vw_ReLeasedUnits_LastTitleHistory "
    End Sub
Public Sub SetWhere()
    SQLz = ""
    SQLz = SQLz & " WHERE Not AssetID is null "
    If Me.AssetSearch & "test" <> "test" Then SQLz = SQLz & " AND AssetID=" & Me.AssetSearch & " "
    If Me.BankSearch & "test" <> "test" Then SQLz = SQLz & " AND BankID=" & Me.BankSearch & " "
    If Me.ClientSearch & "test" <> "test" Then SQLz = SQLz & "AND ClientID=" & Me.ClientSearch & " "
    If Me.TitleNumSearch & "test" <> "test" Then SQLz = SQLz & "AND TitleNumber='" & Me.TitleNumSearch & "' "
    If Me.VINSearch & "test" <> "test" Then SQLz = SQLz & " AND VIN Like'*" & Me.VINSearch & "*' "
    If IsDate(Me.StartDate) And IsDate(Me.EndDate) Then
        SQLz = SQLz & "AND  LastDate >=#" & Me.StartDate & "# AND  LastDate <=#" & Me.EndDate & "# "
        Header = Header & " and By Date Range: " & Me.StartDate & " ~ " & Me.EndDate
    End If
    SQLz = SQLz & " ORDER BY VIN, AssetID"
End Sub


Private Sub OpenSchedule_Click()
    If IsNull(Me.AssetID) Then Exit Sub
    DoCmd.OpenForm "ScheduleControl"
    Forms!ScheduleControl.Form.SearchAssetID2.SetFocus
    Forms!ScheduleControl.Form.SearchAssetID2.Text = Me.AssetID
    Forms!ScheduleControl.Form.UnitLookUp = Me.AssetID
    Form_ScheduleControl.Remotelookup

End Sub

Private Sub OpenUnit_Click()
    If IsNull(Me.AssetID) Then Exit Sub
    Call OpenUnitsForm(Me.AssetID)
End Sub

Private Sub ClientSearch_Change()
    NeedSearch = 1
End Sub
Private Sub AssetSearch_Change()
    NeedSearch = 1
End Sub
Private Sub BankSearch_Change()
    NeedSearch = 1
End Sub
Private Sub TitleNumSearch_Change()
    NeedSearch = 1
End Sub
Private Sub EndDate_Change()
    NeedSearch = 1
End Sub
Private Sub StartDate_Change()
    NeedSearch = 1
End Sub
Private Sub VINSearch_Change()
    NeedSearch = 1
End Sub
