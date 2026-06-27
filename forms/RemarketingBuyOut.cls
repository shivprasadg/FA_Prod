Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim RST As DAO.Recordset
Public Sub ClearBoxes()
Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
     Next Ccrl
End Sub

Private Sub ClearClientSearch_Click()
Me.SearchClient = ""
Call SearchClient_Change
End Sub

Private Sub ClearFilters_Click()
Me.ReMarketingBuyOut_subform.Form.FilterOn = False
End Sub

Private Sub CloseMe_Click()
DoCmd.OpenForm "FAMainMenu"
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
           "From vw_SixKeys " & _
           "GROUP BY clientID, ClientShortName, CompanyName " & _
           "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
    'Me.SearchAsset = ""
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    Call DelWorkTable
    Call ClearBoxes
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.ReMarketingBuyOut_subform.Visible = False
    Me.UnitCount.Visible = False
End Sub

Private Sub GetAll_Click()
    For i = 0 To Units_List.ListCount - 1
        Units_List.Selected(i) = True
    Next i
    Call GetSelectedUnits_Click
    Call GetUnitCounts
End Sub

Private Sub GetSelectedUnits_Click()
    Dim ItemIndex As Variant
    On Error GoTo GetSelectedUnits_Click_Error

    Call DelWorkTable
    For Each ItemIndex In Me.Units_List.ItemsSelected
        SQLz = "INSERT INTO ReMarketingBuyOut_WorkTable ( UnitID ) SELECT " & Me.Units_List.ItemData(ItemIndex) & " AS UnitID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next
    ReMarketingBuyOut_subform.Visible = True

    Me.UnitCount.Visible = True
    Me.ReMarketingBuyOut_subform.Visible = True
    Me!ReMarketingBuyOut_subform.Form.ClearBoxes
    Me.ReMarketingBuyOut_subform.Requery
       Call ClearBoxes
    Exit Sub
fixme1:

    On Error GoTo 0
    Exit Sub

GetSelectedUnits_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSelectedUnits_Click of VBA Document Form_RemarketingBuyOut"
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From ReMarketingBuyOut_WorkTable;")
    DoCmd.SetWarnings True
    ' Me!BulkChange_ChildParts.Form.Detail.Visible = False
End Sub
Private Sub ReQueryAll()
    Me.ReMarketingBuyOut_subform.Requery
End Sub

Private Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys " & _
    "GROUP BY MLID, MLNo, MLOrig, ClientID Having ClientID =" & Me.ClientList & " " & _
    "ORDER BY MLOrig;"
    Me.LeaseList.RowSource = SQLz
    Me.ScheduleList.RowSource = ""
    Me.SchID = ""
    'Me.ReMarketingBuyOut_subform.Form.RecordSource = ""
    Me.Units_List.RowSource = ""
    Me.ReMarketingBuyOut_subform.Visible = False
     Call GetUnitCounts
End Sub

Private Sub Leaselist_Click()
   SQLz = "SELECT SchID, Schedule, MLID, Count(UnitId) As Units, ClientShortName, MLNo, MLOrig " & _
          "FROM vw_SixKeys " & _
          "WHERE MLID = " & Me.LeaseList & " " & _
          "GROUP By SchID, Schedule, MLID, ClientShortName, MLNo, MLOrig Order By CDbl(NumInString(left(Schedule,3))) Desc, Schedule Desc "
    Me.ScheduleList.RowSource = SQLz
    Me.SchID = ""
    'Me.ReMarketingBuyOut_subform.Form.RecordSource = ""
    Me.Units_List.RowSource = ""
    ReMarketingBuyOut_subform.Visible = False
     Call GetUnitCounts
End Sub

Private Sub GetUnitCounts()
Dim UnitX, UnitZ As String
If ([Units_List].[ListCount] - 1) = 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
Me.GetAll.caption = ChrW(9660) & "All " & ([Units_List].ListCount - 1) & " Listed " & UnitX
If ([Units_List].[ListCount] - 1) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
Me.GetSelectedUnits.caption = ChrW(9660) & "Selected (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ
If [Units_List].[ItemsSelected].[Count] > 0 Then Me.GetSelectedUnits.Enabled = True Else Me.GetSelectedUnits.Enabled = False
End Sub
Private Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.SchNo = Me.ScheduleList.Column(1)
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Me.ReMarketingBuyOut_subform.Visible = False
    SQLz = "SELECT AssigneeBanks.* From AssigneeBanks WHERE AssigneeBanks.SchID=" & Me.ScheduleList & ";"
    Me.BankList.RowSource = SQLz
    'Me.SearchAsset = ""
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    Call GetUnitCounts
End Sub

'Private Sub SearchAsset_Change()
'   ' Me.SearchAsset = ""
'    Me.SearchUnit = ""
'    Me.SearchVin = ""
'    Me.SearchStatus = ""
'    Me.SearchBank = ""
'    Me.SearchLoan = ""
'    SQLz = "SELECT * " & _
'           "FROM RMKTScheduleGroupUnit " & _
'           "Where SchID=" & Me.SchID & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' Order BY UnitNum;"
'    Me.Units_List.RowSource = SQLz
'    Call GetUnitCounts
'End Sub



Private Sub SearchClient_Change()
SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
       "From vw_SixKeys WHERE ClientShortName Like '*" & Me.SearchClient.Text & "*' Or CompanyName Like '*" & Me.SearchClient.Text & "*' " & _
       "GROUP BY clientID, ClientShortName, CompanyName " & _
       "ORDER BY CompanyName;"

 'Sqlz = "SELECT Client, clientID FROM ClientsList " & _
           "WHERE Client Like '*" & Me.SearchClient.Text & "*' Or ClientShNm Like '*" & Me.SearchClient.Text & "*' ORDER BY Client;"
    Me.ClientList.RowSource = SQLz
End Sub

Private Sub SearchBuyOut_Change()
    'Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " AND BuyOutDate Like '*" & Me.SearchBuyOut.Text & "*' Order BY BuyOutDate;"
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub



Private Sub SearchUnit_Change()
    Me.SearchBuyOut = ""
    'Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " AND UnitNum Like '*" & Me.SearchUnit.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLz
     Call GetUnitCounts
End Sub

Private Sub SearchVIN_Change()
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    'Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " and VIN Like '*" & Me.SearchVin.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLz
     Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    'Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " and S Like '*" & Me.SearchStatus.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLz
     Call GetUnitCounts
End Sub
Private Sub SearchBank_Change()
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    'Me.SearchBank = ""
    Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " and Bank Like '*" & Me.SearchBank.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLz
     Call GetUnitCounts
End Sub

Private Sub SearchLoan_Change()
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    'Me.SearchLoan = ""
    SQLz = "SELECT * " & _
           "FROM RMKTScheduleGroupUnit " & _
           "Where SchID=" & Me.SchID & " and [Loan#] Like '*" & Me.SearchLoan.Text & "*' Order BY UnitNum;"
    Me.Units_List.RowSource = SQLz
     Call GetUnitCounts
End Sub
Private Sub ClearLabel_Click()
    Me.SearchBuyOut = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchBank = ""
    Me.SearchLoan = ""
    Me.SearchVin.SetFocus
   Call SearchVIN_Change
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

Private Sub ToggleSortBank_Click()
Call SortListBox("ToggleSortStatus", "Units_List", "Bank")
End Sub

Private Sub ToggleSortBuyOutDate_Click()
Call SortListBox("ToggleSortStatus", "Units_List", "BuyOutDate")
End Sub

Private Sub ToggleSortLoan_Click()
Call SortListBox("ToggleSortStatus", "Units_List", "Loan#")
End Sub

Private Sub ToggleSortStatus_Click()
    Call SortListBox("ToggleSortStatus", "Units_List", "S")
End Sub

Private Sub ToggleSortUnit_Click()
Call SortListBox("ToggleSortStatus", "Units_List", "UnitNum")
End Sub

Private Sub ToggleSortVIN_Click()
 Call SortListBox("ToggleSortStatus", "Units_List", "VIN")
End Sub

Private Sub Units_List_Click()
 Call GetUnitCounts
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
On Error GoTo FixMe
  Set RST = Me.ReMarketingBuyOut_subform.Form.RecordsetClone
  RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
  Me.ReMarketingBuyOut_subform.Form.Bookmark = RST.Bookmark
  Set rs = Nothing
FixMe:
End Sub
