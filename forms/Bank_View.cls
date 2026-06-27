Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz As String
Dim SQLx As String
Private Sub AddContactBtn_Click()
    DoCmd.OpenForm "contact_view", , , , acFormAdd, acDialog
End Sub

Private Sub AddContactBtn_DblClick(Cancel As Integer)
    DoCmd.OpenForm "contact_view", , , , acFormAdd, acDialog
End Sub

Private Sub AddContactBtnFA_Click()
    DoCmd.OpenForm "clientRef", acNormal, , "clientRef = " & Me.BankId, acFormPropertySettings, acDialog
End Sub

Private Sub AddContactBtnFA_DblClick(Cancel As Integer)
    DoCmd.OpenForm "contact_view", acNormal, , "clientRef = " & Me.BankId, acFormPropertySettings, acDialog
End Sub


Private Sub AddLocBtn_Click()
    DoCmd.OpenForm "Locations_view3", , , , acFormAdd, acDialog
End Sub

Private Sub AddNewCompany_Click()
    On Error GoTo Err_AddNewCompany_Click
    DoCmd.GotoRecord , , acNewRec
Exit_AddNewCompany_Click:
    Exit Sub
Err_AddNewCompany_Click:
    MsgBox Err.Description
    Resume Exit_AddNewCompany_Click

End Sub

Private Sub AllBankExcel_Click()
    FormatExcel "BankListAll", "C:\Faas\" & FAUserName & "\Temp\BankListAll.xlsx", "Banks", "yes", , 0, "All Banks", False
End Sub

Private Sub AllBankList_Click()
    DoCmd.OpenReport "AllBankList", acViewReport
End Sub

Private Sub BankList_Click()
    Me.Form.Filter = ""
    Me.Form.FilterOn = False
    Set RST = Me.RecordsetClone
    RST.FindFirst "BankId = " & Me.BankList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    Set RST = Nothing
    Forms!Bank_View!Bank_Locations.Form.SetLocations
    Me.PickModeToggle = False
    Call PickModeToggle_Click
    Me.Units_List.Requery
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub EditContactsBtn_Click()
    DoCmd.OpenForm "contact_view", acNormal, , clientRef = " & Me.BankId, acFormPropertySettings, acDialog"
End Sub

Private Sub DrawDownMaker_Click()
    DoCmd.close acForm, "DrawDownMaker"
    DoCmd.OpenForm "DrawDownMaker"
    Forms!DrawDownMaker.BankList = Me.BankId
    Form_DrawDownMaker.BankList_Click
End Sub

Private Sub DrawDownManager_Click()
    DoCmd.close acForm, "DrawDownManager"
    DoCmd.OpenForm "DrawDownManager"
End Sub

Private Sub EditContactsBtnBnk_Click()
    DoCmd.OpenForm "bankcontact_view", acNormal, , "clientRef = " & Me.BankId, acFormPropertySettings, acDialog
End Sub

Private Sub EditLocBtn_Click()
    DoCmd.OpenForm "Location form", acNormal, , "clientRef = " & Me.BankId, acFormPropertySettings, acDialog
End Sub

Private Sub Form_Current()
    Dim RAddr  As String
    If DCount("LocationID", "Locations", "LocationType='Billing' AND ClientREF=" & Me.BankId) > 0 Then
        SQLz = "Select * From Locations Where ClientRef=" & Nz(Me.BankId, 0) & " AND LocationType='Billing';"
        Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        RST.MoveFirst
        RAddr = RST!LocationStreet & vbNewLine
        If Nz(RST!LAddrLine1, "") <> "" Then RAddr = RAddr & RST!LAddrLine1 & vbNewLine
        If Nz(RST!LAddrLine2, "") <> "" Then RAddr = RAddr & RST!LAddrLine2 & vbNewLine
        RAddr = RAddr & RST!LCity & " " & RST!LState & " " & RST!LZip
    Else
        RAddr = "No Remit To Address Set!"
    End If
    Me.RemitAddress = RAddr
    RST.close
    Set RST = Nothing
    Me.SearchAssetId = ""
    Me.SearchVin = ""
    Me.SearchUnitNum = ""
    Me.SearchSchedule = ""
    Me.SearchClient = ""
End Sub

Private Sub Form_Load()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From Pick_Units_WorkTable;")
    DoCmd.SetWarnings True
End Sub

Private Sub GetAll_Click()
    For i = 0 To Units_List.ListCount - 1
        Units_List.Selected(i) = True
    Next i
    Call GetSelectedUnits_Click
End Sub

Private Sub GetSelectedUnits_Click()
    Call SetUnits
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From Bulk_Update_WorkTable;")
    DoCmd.RunSQL ("DELETE * From Bulk_Update_WorkTable_Child;")
    DoCmd.RunSQL ("DELETE * From UnitWorkTable;")
    DoCmd.RunSQL ("DELETE * From UnitTitleLog_Work;")
    DoCmd.RunSQL ("DELETE * From Pick_Units_WorkTable;")
    DoCmd.SetWarnings True
    Me!BulkChange_ChildParts.Form.Detail.Visible = False
    ' Call ReQueryAll
End Sub
Public Function SetUnits(Optional fFilter As String)
    Dim ItemIndex As Variant
    Dim UnitIDx As String
    Call DelWorkTable
    For Each ItemIndex In Me.Units_List.ItemsSelected
        UnitIDx = UnitIDx & Me.Units_List.ItemData(ItemIndex) & ","
    Next
    If Right(UnitIDx, 1) = "," Then UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
    ' Make Mass-Mod Work table.
    SQLx = "INSERT INTO Bulk_Update_WorkTable (UnitID, UnitDesc, ChildPartCost) " & _
           "Select UnitID, [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitModel] & ' ' & [UnitSubType] & ' ' & [UnitType] " & _
           ",nz(CPC,0) From Units " & _
           "Left Join (Select UnitREf, Sum(childFinalCost) as CPC From ChildParts " & _
           "Group by UnitREf) As CPCost ON Units.UnitID = CPCost.UnitRef " & _
           "Where UnitID In (" & UnitIDx & ")"
    If fFilter <> "" Then SQLx = SQLx & " AND " & fFilter & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLx)
    ' Make Title Log Work table for new Mass-Mod Titles.
    SQLx = ""
    SQLx = "INSERT INTO UnitTitleLog_Work ( AssetID, VIN, TitleNumber, TitleSt, [Desc], LastDirection, UserID, LastEntry, LastEntryTypeId, UnitStatus) "
    SQLx = SQLx & "SELECT S.AssetID, S.VIN, H.TitleNumber, S.TitledState, S.UnitDescShortVer"
    SQLx = SQLx & ", IIf([EntryTypeId]=1,'IN',IIf([EntryTypeID]=2,'Out','-')) AS Dir, '" & FAUserName & "' AS [User]"
    SQLx = SQLx & ", nz(H.EntryReason & ' (' & H.UnitTitleLogDate &')','N/A')"
    SQLx = SQLx & ",EntryTypeId, [Status] "
    SQLx = SQLx & "FROM (vw_SixKeys AS S INNER JOIN Units AS U ON S.AssetID = U.UnitID) "
    SQLx = SQLx & "LEFT JOIN UnitTitleHistoryLastEntry AS H ON U.UnitID = H.AssetID "
    SQLx = SQLx & "WHERE (((S.AssetID) In (Select UnitID From Bulk_Update_WorkTable)));"
    DoCmd.RunSQL (SQLx)

    SQLx = ""
    DoCmd.SetWarnings True
    ''    Me.clientID = Me.ClientList.Column(0)
    Call ReQueryAll
    On Error GoTo fixme1
    Me.UnitCount.Visible = True
    Me.UnitCount = Me.BulkChange_Desc.Form!UnitCount
    Exit Function
fixme1:
End Function
Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount]) > 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] > 0 Then UnitSelected = [Units_List].[ItemsSelected].[Count] Else UnitSelected = 0
    Me.GetAll.caption = ChrW(9660) & "All " & ([Units_List].ListCount) & ""    '& UnitX
    If ([Units_List].[ListCount]) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
    Me.GetSelectedUnits.caption = ChrW(9660) & "Selected (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ
    If [Units_List].[ItemsSelected].[Count] > 0 Then Me.GetSelectedUnits.Enabled = True Else Me.GetSelectedUnits.Enabled = False
End Sub

Private Sub MaketitleRequest_Click()
If FAUserFullName = "Leo Johnson" Then Call RunTitleRequest(True) Else Call RunTitleRequest(False)        '' Change to True for Batch Run
End Sub
Public Function RunTitleRequest(Optional ITBatch As Boolean)
If ITBatch = True Then GoTo BatchRun
    If Me.PickModeToggle = False Then
        Dim ItemVar As Variant
        For Each ItemVar In Me.Units_List.ItemsSelected
            SQLz = "Insert into Pick_Units_WorkTable(UnitID) SELECT " & Me.Units_List.ItemData(ItemVar) & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
        Next ItemVar
    End If
    If Nz(Me.UnitSelected, 0) = 0 Or DCount("UnitID", "Pick_Units_WorkTable") = 0 Then
        MsgBox "You haven't Selected any units" & vbNewLine & "(is the 'Pick On' Button pressed?)", vbCritical, "No units, No service"
        Exit Function
    End If
BatchRun:
    If Me.RequestType = 7 And Me.OtherReason = "" Then
        MsgBox "You Selected Reason 7 (Other) but you do not have the anything entered in for the 'Other' reason", "You're silly"
        Exit Function
    End If
    Dim Otherx As String
    If Me.RequestType = 7 Then Otherx = Me.OtherReason
    Dim Name2Use As String
    If FAUserFullName = "Leo Johnson" Then Name2Use = "Ysis Jackson" Else Name2Use = FAUserFullName
    Call MakeTitleRequestLetter(Me.BankId, Me.RequestType, Otherx, , Name2Use)
End Function

Private Sub PickModeToggle_Click()
    Dim i As Integer
    For i = Me.Units_List.ListCount - 1 To 0 Step -1
        Me.Units_List.Selected(i) = 0
    Next i
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Pick_Units_WorkTable")
    DoCmd.SetWarnings True
    If Me.PickModeToggle = True Then
        Me.PickModeToggle.caption = ChrW(9660) & "Pick On"
    Else
        Me.PickModeToggle.caption = "Pick Off"
    End If
    Call GetUnitCounts
End Sub

Public Function SearchList(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetId = ""
    If sField <> "VIN" Then Me.SearchVin = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum = ""
    If sField <> "Schedule" Then Me.SearchSchedule = ""
    If sField <> "Client" Then Me.SearchClient = ""
    If Len(Me("Search" & sField).Text) >= 1 And Len(Me("Search" & sField).Text) < 3 Then Exit Function
    SQLz = "SELECT AssetID, Vin, UnitNum,uStatus, Client, Schedule & '|' & Group, GroupID " & _
        "From UnitTitleHistoryLastEntry " & _
        "WHERE (EntryReasonId=1 or EntryReasonId=4 or  EntryReasonId=25 or EntryReasonId=36) AND BankID =" & Me.BankId & " AND " & sField & " Like '*" & Me("Search" & sField).Text & "*' " & _
        "Order By " & sField & ""
    Me.Units_List.RowSource = SQLz
    Call updateGetall
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
    Call GetUnitCounts
End Function

Private Sub RequestType_Click()
    If Me.RequestType = 7 Then
    Me.OtherReason.Enabled = True
    Else
        Me.OtherReason.Enabled = False
        Me.OtherReason = ""
    End If
End Sub

Private Sub SearchAssetId_Change()
    Call SearchList("AssetID")
End Sub
Private Sub SearchClient_Change()
    Call SearchList("Client")
End Sub
Private Sub SearchSchedule_Change()
    Call SearchList("Schedule")
End Sub
Private Sub SearchUnitNum_Change()
    Call SearchList("UnitNum")
End Sub
Private Sub SearchVIN_Change()
    Call SearchList("VIN")
End Sub

Private Sub SearchAssetID_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchSchedule_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchClient_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchUnitNum_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchVIN_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub

Private Sub updateGetall()
    Me.GetAll.caption = ChrW(9660) & "All " & [Units_List].ListCount
End Sub

Public Function ToggleSort(toggleField As String)
    If toggleField <> "AssetID" Then Me.ToggleSortAssetID = 0
    If toggleField <> "VIN" Then Me.ToggleSortVIN = 0
    If toggleField <> "UnitNum" Then Me.ToggleSortUnitNum = 0
    If toggleField <> "Client" Then Me.ToggleSortClient = 0
    If toggleField <> "Schedule" Then Me.ToggleSortSchedule = 0
    If toggleField <> "uStatus" Then Me.ToggleSortuStatus = 0
    SQLz = SplitString(Units_List.RowSource, " Order By", 0) & " Order By " & toggleField & " "
    If Me("ToggleSort" & toggleField) = 0 Then SQLz = SQLz & " Desc"
    Me.Units_List.RowSource = SQLz
End Function
Private Sub ToggleSortAssetID_Click()
    Call ToggleSort("AssetID")
End Sub
Private Sub ToggleSortClient_Click()
    Call ToggleSort("Client")
End Sub
Private Sub ToggleSortUnitNum_Click()
    Call ToggleSort("UnitNum")
End Sub
Private Sub ToggleSortSchedule_Click()
    Call ToggleSort("Schedule")
End Sub
Private Sub ToggleSortuStatus_Click()
    Call ToggleSort("uStatus")
End Sub
Private Sub ToggleSortVIN_Click()
    Call ToggleSort("VIN")
End Sub

Private Sub Units_List_Click()
    If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.Units_List.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Pick_Units_WorkTable(UnitID) SELECT " & Me.Units_List.Column(0) & ";")
        DoCmd.SetWarnings True
        Call ReselectPickedUnits
    End If
    Call GetUnitCounts
End Sub
Public Sub ReselectPickedUnits()
    Dim i, x As Integer
    Dim pUnits As String
    SQLz = "Select * From Pick_Units_WorkTable Order by UnitID Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For x = 1 To RST.RecordCount
            For i = Me.Units_List.ListCount - 1 To 0 Step -1
                If Me.Units_List.ItemData(i) = CStr(RST!UnitID) Then Me.Units_List.Selected(i) = 1
            Next i
            RST.MoveNext
        Next x
    End If
    RST.close
    Set RST = Nothing
End Sub
