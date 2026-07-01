Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, TargetWhere As String

Private Sub ClearSeachBoxText_Click()
  Call ClearSearches
End Sub
Private Sub ClearSeachBoxText2_Click()
Call ClearSearches
End Sub

Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExcelReport_Click()
Dim XLS, Header As String
Header = "Fleet List " & Date
XLS = "|F,TX|I,BC|W,$E|"
SendReport2Excel Me.UnitsList.RowSource, "Fleet List", Header, True, XLS, Me.UnitsList.ListCount, True
End Sub

Private Sub Form_Load()
  ' Me.UnitSelection = 1
  ' TargetWhere = " Where MLOrig <> '' "
  SQLz = "SELECT ClientShortName AS Client FROM vw_SixKeys GROUP BY ClientShortName Order By ClientShortName;"
  Me.ClientList.RowSource = SQLz
  ' Call ResetForm_Click
End Sub
Private Sub ClearSearches(Optional SearchBox As String)

  Dim Ccrl  As Control
  For Each Ccrl In Me.Controls
    If Me(Ccrl.Name).Tag = "SearchX" Then
      If Not IsNull(SearchBox) Then
        If SearchBox <> Me(Ccrl.Name).Name Then Me(Ccrl.Name) = ""
      Else
        Me(Ccrl.Name) = ""
      End If
    End If
  Next Ccrl
  Me.UnitsList = ""
End Sub

Private Sub OpenBuyer_Click()
    If Me.BuyerID = "" Or IsNull(Me.BuyerID) Then Exit Sub
    Call OpenBuyersFrom(DLookup("BuyerID", "vw_BuyersOrder", "UnitID=" & Me.AssetID & ""), Nz(DLookup("InvoiceRequestID", "vw_BuyersOrder", "UnitID=" & Me.AssetID & ""), 0))
End Sub

Private Sub OpenClient_Click()
  If Me.ClientID = "" Or IsNull(Me.ClientID) Then Exit Sub
  DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
  DoCmd.SelectObject acForm, "ClientsView"
End Sub

Private Sub OpenGroup_Click()
  If Me.GroupID = "" Or IsNull(Me.GroupID) Then Exit Sub
  DoCmd.OpenForm "GroupForm", , , "SGrpid=" & Me.GroupID
  DoCmd.SelectObject acForm, "GroupForm"
End Sub

Private Sub OpenMasterLease_Click()
  If Me.ClientID = "" Or IsNull(Me.ClientID) Then Exit Sub
  DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
  DoCmd.SelectObject acForm, "ClientsView"
  Forms!ClientsView.TabControls.Value = 7
End Sub

Private Sub OpenSchedule_Click()
  If Me.SchID = "" Or IsNull(Me.SchID) Then Exit Sub
  DoCmd.OpenForm "Schedule_Form", , , "Schid=" & Me.SchID
  DoCmd.SelectObject acForm, "Schedule_Form"
End Sub

Private Sub OpenScheduleControl_Click()
  If Me.AssetID = "" Or IsNull(Me.AssetID) Then Exit Sub
  DoCmd.OpenForm "ScheduleControl"
  Forms!ScheduleControl.Form.SearchAssetID2.SetFocus
  Forms!ScheduleControl.Form.SearchAssetID2.Text = Me.AssetID
  Forms!ScheduleControl.Form.UnitLookUp = Me.AssetID
  Form_ScheduleControl.Remotelookup
  DoCmd.SelectObject acForm, "ScheduleControl"
End Sub

Private Sub OpenUnit_Click()
  If Me.AssetID = "" Or IsNull(Me.AssetID) Then Exit Sub
  Call OpenUnitsForm(Me.AssetID)
  'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.AssetID
  Me.SearchAsset.SetFocus
  DoCmd.SelectObject acForm, "UnitsForm"
End Sub

Private Sub ResetForm_Click()
  Call ClearSearches
  SQLz = "Select * From vw_UnitReverseLookUp Where MlOrig is not null Order By Client;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchALL_Button_Click()
'If Me.SearchClient & Me.SearchLocation & Me.SearchMaster & Me.SearchSchedule & Me.SearchGroup & Me.SearchAsset & SearchUnitNum & Me.SearchStatus & SearchVin & Me.SearchDesc & Me.SearchBuyer <> "" Then
  'Me.UnitSelection = 1
  TargetWhere = " Where AssetID > 0 "
  SQLz = "Select * From vw_UnitReverseLookUp " & TargetWhere & " "
  If Me.SearchClient <> "" Then SQLz = SQLz & "AND Client Like '*" & Me.SearchClient & "*' "
  If Me.SearchFAID <> "" Then SQLz = SQLz & "AND FAID Like '*" & Me.SearchFAID & "*' "
  If Me.SearchMaster <> "" Then SQLz = SQLz & "AND MLA Like '*" & Me.SearchMaster & "*' "
  If Me.SearchSchedule <> "" Then SQLz = SQLz & "AND Schedule Like '*" & Me.SearchSchedule & "*' "
  If Me.SearchGroup <> "" Then SQLz = SQLz & "AND UnitGroup Like'*" & Me.SearchGroup & "*' "
  If Me.SearchAsset <> "" Then SQLz = SQLz & "AND AssetID Like '*" & Me.SearchAsset & "*' "
  If Me.SearchUnitNum <> "" Then SQLz = SQLz & "AND [Unit#] Like '*" & Me.SearchUnitNum & "*' "
  If Me.SearchStatus <> "" And Me.SearchStatus Like "-*" Then SQLz = SQLz & "AND NOT Status Like '*" & SplitString(Me.SearchStatus, "-", 1) & "*' "
  If Me.SearchStatus <> "" And Not Me.SearchStatus Like "-*" Then SQLz = SQLz & "AND Status Like '*" & Me.SearchStatus & "*' "
  If Me.SearchVin <> "" Then SQLz = SQLz & "AND VIN Like '*" & Me.SearchVin & "*' "
  If Me.SearchDesc <> "" Then SQLz = SQLz & "AND UnitDesc Like '*" & Me.SearchDesc & "*' "
  If Me.SearchBuyer <> "" Then SQLz = SQLz & "AND Buyer Like '*" & Me.SearchBuyer & "*' "
  If Me.SearchLocation <> "" Then SQLz = SQLz & "AND Location Like '*" & Me.SearchLocation & "*' "
  If Me.SearchPlate <> "" Then SQLz = SQLz & "AND [Plate#] Like '*" & Me.SearchPlate & "*' "
  SQLz = SQLz & "Order By Client, SchID, SGrpID, UnitDesc desc, AssetID;"
  Me.UnitsList.RowSource = SQLz
  ' End If
End Sub

Private Sub SearchAsset_Button_Click()
  Call ClearSearches("SearchAsset")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE AssetID Like'*" & Me.SearchAsset & "*' Order By AssetID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchBuyer_Button_Click()
  Call ClearSearches("SearchBuyer")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE Buyer Like'*" & Me.SearchBuyer & "*' Order By Buyer, Client, UnitDesc desc;"
  Me.UnitsList.RowSource = SQLz
End Sub
Private Sub ClientList_Click()
Me.SearchClient = Me.ClientList
Call SearchClient_Button_Click
End Sub
Private Sub SearchClient_Button_Click()
  Call ClearSearches("SearchClient")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE Client Like'*" & Me.SearchClient & "*' Order By Client, UnitDesc;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchDesc_Button_Click()
  Call ClearSearches("SearchDesc")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE UnitDesc Like'*" & Me.SearchDesc & "*' Order By UnitDesc;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchFAID_Button_Click()
  Call ClearSearches("SearchFAID")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE FAID Like'*" & Me.SearchFAID & "*' Order By FAID, Client;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchGroup_Button_Click()
  Call ClearSearches("SearchGroup")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE UnitGroup Like'*" & Me.SearchGroup & "*' Order By Client, SchID, SGrpID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchLvvvocation_Button_Click()
  Call ClearSearches("SearchLocation")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE Location Like'*" & Me.SearchLocation & "*' Order By SearchLocation, Client, SchID, SGrpID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchMaster_Button_Click()
  Call ClearSearches("SearchMaster")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE MLA Like'*" & Me.SearchMaster & "*' Order By Client, SchID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchPlate_Button_Click()
  Call ClearSearches("SearchPlate")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE [Plate#] Like'*" & Me.SearchPlate & "*' Order By [Plate#], AssetID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchStatus_Button_Click()
  Call ClearSearches("SearchStatus")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE "
  If Me.SearchStatus Like "*,*" Then
  Dim ii As Integer

  End If
  If Me.SearchStatus Like "-*" Then SQLz = SQLz & " NOT Status Like '*" & SplitString(Me.SearchStatus, "-", 1) & "*' " _
  Else SQLz = SQLz & "Status Like '*" & Me.SearchStatus & "*' "
  SQLz = SQLz & "Order By Client, SchID, SGrpID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchUnitNum_Button_Click()
  Call ClearSearches("SearchUnitNum")
  SQLz = "Select * From vw_UnitReverseLookUp Where [Unit#] Like'*" & Me.SearchUnitNum & "*' Order By [Unit#];"
  Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchSchedule_Button_Click()
  Call ClearSearches("SearchSchedule")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE Schedule Like '*" & Me.SearchSchedule & "*' Order By Client, SchID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub SearchVIN_Button_Click()
  Call ClearSearches("SearchVIN")
  SQLz = "Select * From vw_UnitReverseLookUp WHERE VIN Like '*" & Me.SearchVin & "*' Order By Vin, SchID, SGrpID;"
  Me.UnitsList.RowSource = SQLz
End Sub

Private Sub UnitSelection_Click()
  If Me.UnitSelection = 1 Then
    TargetWhere = " Where MLOrig Is Not Null "
    SQLz = "SELECT ClientShortName AS Client FROM vw_SixKeys GROUP BY ClientShortName Order By ClientShortName;"
  End If
  If Me.UnitSelection = 2 Then
    TargetWhere = " Where MLOrig = 'FA' "
    SQLz = "SELECT ClientShortName AS Client FROM vw_SixKeys WHERE MLOrig ='FA' GROUP BY ClientShortName Order By ClientShortName;"
  End If
  If Me.UnitSelection = 3 Then
    TargetWhere = " Where MLOrig <> 'FA' "
    SQLz = "SELECT ClientShortName AS Client FROM vw_SixKeys WHERE MLOrig <>'FA' GROUP BY ClientShortName Order By ClientShortName;"
  End If
  Me.ClientList.RowSource = SQLz
  Call ResetForm_Click
End Sub

Private Sub UnitsList_Click()
  Me.AssetID = Me.UnitsList.Column(5)
  Me.ClientID = Me.UnitsList.Column(15)
  Me.MLID = Me.UnitsList.Column(16)
  Me.SchID = Me.UnitsList.Column(17)
  Me.GroupID = Me.UnitsList.Column(18)
  Me.LocationID = Me.UnitsList.Column(19)
  Me.BuyerID = Me.UnitsList.Column(20)
End Sub

Private Sub UnitsList_DblClick(Cancel As Integer)
  If Me.AssetID & "test" <> "test" Then Call OpenUnit_Click
End Sub
