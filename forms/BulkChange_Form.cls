Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx, SQLz, ReportX As String

Private Sub Editdraw_Click()
     If Me.DrawDownList & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "DrawDownAddaDraw"
    Call DoCmd.OpenForm("DrawDownAddaDraw", , , , , , "Edit>" & Me.DrawDownList)
End Sub

Private Sub Form_Load()
    Call ClearBoxes
    Call DelWorkTable
    Call ClearLists("All")
    Call SearchUnits_List("ClearAll")
    Call ToggleSort("ClearAll")
    Me.GrpID = ""
    Me.ModMade = 0
    Me.PickModeToggle = False
    Call PickModeToggle_Click
    Me.ClientList.Requery
End Sub

Public Sub ClientList_Click()
    If Nz(Me.ScheduleList, 0) <> 0 Then
        Call CheckForReasons
        Call RePairExchangeIT
    End If
    SQLz = "SELECT MLID, MLNo, MLOrig From vw_SixKeys " & _
           "WHERE ClientId =" & Me.ClientList.Column(0) & " GROUP BY MLID, MLNo, MLOrig Order by MLOrig, MLNo;"
    Me.LeaseList.RowSource = SQLz
    Me.ClientID = Me.ClientList.Column(0)
    Me.ClientGroupID = Me.ClientList.Column(3)
    Me.GetSelectedUnits.Visible = False
    Me.GetAll.Visible = False
    Me.GroupInstructions.Visible = True
    Me.UnitCount.Visible = False
    Me.GrpID = ""
    Me.SchNo = Me.ClientList.Column(1)
    Call ClearLists("Client,Lease")
    Call DelWorkTable
    Call SearchUnits_List("ClearAll")
    Call ToggleSort("ClearAll")
    Me.PickModeToggle = False
    Call PickModeToggle_Click
End Sub

Public Sub Leaselist_Click()
    SQLz = "SELECT SchID, Schedule, Count(AssetID) AS Units From vw_SixKeys " & _
           "WHERE MLID =" & Me.LeaseList.Column(0) & " GROUP BY SchID, Schedule Order by CDbl(NumInString(Left(Schedule,3))) Desc, Schedule Desc;"
    Me.ScheduleList.RowSource = SQLz
    Me.UnitCount.Visible = False
    Me.GetSelectedUnits.Visible = False
    Me.GetAll.Visible = False
    Me.MLID = Me.LeaseList.Column(0)
    Me.GrpID = ""
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1)
    If Me.LeaseList.Column(2) = "FA" Then
        Call CheckForReasons
        Call RePairExchangeIT
    End If
    Me.PickModeToggle = False
    Call ClearLists("Client,Lease,Schedule")
    Call DelWorkTable
    Call SearchUnits_List("ClearAll")
    Call ToggleSort("ClearAll")
    Call PickModeToggle_Click
End Sub

Private Sub OpenGroupForm_Click()
    If Me.GrpID & 0 = 0 Then Exit Sub
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.GrpID
End Sub

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

Public Sub ScheduleList_Click()
    If Me.GrpID & "test" <> "test" Then
        If DMax("Status", "vw_SixKeys", "SGrpID=" & Me.GrpID) = "A" Then Call AllUnitsInGroup_AcceptedMessage(Me.GrpID)
    End If
    If Me.LeaseList.Column(2) = "FA" Then
        Call CheckForReasons
        Call RePairExchangeIT
    End If
    SQLz = "SELECT SGrpID, UnitGroup, Count(AssetID) AS Units From vw_SixKeys " & _
        "WHERE SchID =" & Me.ScheduleList.Column(0) & " GROUP BY SGrpID, UnitGroup Order by UnitGroup;"
    Me.GroupList.RowSource = SQLz
    Me.GrpID = ""
    Call ClearLists("Client,Lease,Schedule,Group")
    Call DelWorkTable
    Call SearchUnits_List("ClearAll")
    Call ToggleSort("ClearAll")
    If DCount("SGrpID", "SchGrp", "SGrpSchID=" & Me.ScheduleList.Column(0)) > 1 Then
        Me.GetSelectedUnits.Visible = False
        Me.GetAll.Visible = False
        Me.UnitCount.Visible = False
      ' Me.Units_List.Visible = False
    Else
        Me.GroupList = DLookup("SGrpID", "SchGrp", "SGrpSchID=" & Me.ScheduleList.Column(0))
        Me.GrpID = Me.GroupList.Column(0)
        Me.GetSelectedUnits.Visible = True
        Me.GetAll.Visible = True
        Me.UnitCount.Visible = True
        Me.Units_List.Visible = True
    End If
    Me.PickModeToggle = False
    Call PickModeToggle_Click
    Call updateGetall
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1)
End Sub

Private Sub DescList_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub
    SQLz = "SELECT AssetID, UnitVin, UnitNum, UnitDesc, SGrpID From Bulk_Units_FAID_SCH  " & _
        "WHERE SGrpID =" & Me.GroupList.Column(0) & "AND  UnitDesc = '" & DescList.Column(0) & "' Order By AssetID ;"
    Me.Units_List.RowSource = SQLz
    Call updateGetall
    Me.DescList.ControlTipText = DescList.Column(0)
    If Me.LeaseList.Column(2) = "FA" Then
        Call CheckForReasons
        Call RePairExchangeIT
    End If
End Sub
Private Sub DrawDownList_Click()
    If Me.GroupList.Column(0) & "test" = "test" Then Exit Sub

    SQLz = "SELECT AssetID, UnitVin, UnitNum, UnitDesc, SGrpID "
    SQLz = SQLz & "From Bulk_Units_FAID_SCH "
    SQLz = SQLz & "WHERE SGrpID =" & Me.GroupList.Column(0) & " "
    If DrawDownList.Column(0) & "test" = "test" Then
    SQLz = SQLz & "AND IsNULL(DrawDownID) Order By AssetID ;"
    Else
    SQLz = SQLz & "AND DrawDownID = " & DrawDownList.Column(0) & " Order By AssetID ;"
    End If
    Me.Units_List.RowSource = SQLz
    Call updateGetall
  '  Me.DescList.ControlTipText = DescList.Column(0)
    If Me.LeaseList.Column(2) = "FA" Then
        Call CheckForReasons
        Call RePairExchangeIT
    End If
End Sub
Public Sub GroupList_Click()
    If Me.GrpID & "test" <> "test" Then
        If Me.LeaseList.Column(2) = "FA" Then
            Call CheckForReasons
            Call RePairExchangeIT
            Call SetSessionProfile(Me.GroupList.Column(0), "MassMod")
            If DMax("Status", "vw_SixKeys", "SGrpID=" & Me.GrpID) = "A" Then Call AllUnitsInGroup_AcceptedMessage(Me.GrpID)
        End If
    End If

    Me.GrpID = Me.GroupList.Column(0)
    SQLz = "SELECT UnitDesc, Count(AssetID), SGrpID From Bulk_Units_FAID_SCH  " & _
        "WHERE SGrpID =" & Me.GroupList.Column(0) & " Group By  UnitDesc, SGrpID Order By UnitDesc;"
    Me.DescList.RowSource = SQLz

    SQLz = "SELECT U.DrawDownID, D.DrawNumber, Nz(D.DrawName,'None') AS DrawName, Count(U.UnitID) AS Units, U.GroupRef "
    SQLz = SQLz & "FROM Units AS U LEFT JOIN DrawDowns AS D ON U.DrawDownID = D.DrawDownID "
    SQLz = SQLz & "Where (((U.GroupRef) = " & Me.GroupList.Column(0) & ")) "
    SQLz = SQLz & "GROUP BY U.DrawDownID, D.DrawNumber, Nz([d].[DrawName],'None'), U.GroupRef;"
    Me.DrawDownList.RowSource = SQLz

    SQLz = "SELECT AssetID, UnitVin, UnitNum, UnitDesc, SGrpID From Bulk_Units_FAID_SCH  " & _
        "WHERE SGrpID =" & Me.GroupList.Column(0) & " Order By AssetID ;"
    Me.Units_List.RowSource = SQLz

    Me.UnitCount.Visible = False
    Call DelWorkTable
    Call SearchUnits_List("ClearAll")
    Call ToggleSort("ClearAll")
    Me.Units_List.Visible = True
    Me.GetSelectedUnits.Visible = True
    Me.GetAll.Visible = True
    Call updateGetall
    'Me.UnitCount.Visible = True
    Me.GroupInstructions.Visible = True
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1) & " Grp: " & GroupList.Column(1)
    If DLookup("SGRPPLB", "SchGrp", "SGrpId=" & Me.GrpID) = True Then Me.PLBAsOfDate.Visible = True Else Me.PLBAsOfDate.Visible = False
    Me.PLBAsOfDate = DMax("ODOReadDate", "LastMilesReportedAll", "GroupRef=" & Me.GrpID)
    Call GetUnitCounts
    End Sub
Private Sub CheckForReasons()
    Call SyncAuditTables
    If Nz(Me.GrpID, 0) = 0 Or Nz(Me.GrpID, "") = "" Then Exit Sub
    Call CycleForReasons(Me.GrpID)
    If DMax("Status", "vw_SixKeys", "SGrpID=" & Me.GrpID) = "A" Then Call AllUnitsInGroup_AcceptedMessage(Me.GrpID)
    Call MakePaymentAdjustment(Me.GrpID)
    Me.ModMade = 0
End Sub
Private Sub GroupList_DblClick(Cancel As Integer)
    If IsNull(Me.GroupList.Column(0)) Then Exit Sub
    DoCmd.OpenForm "GroupForm", , , "SchID =" & Me.ScheduleList.Column(0) & " "
End Sub

Public Function ClearLists(KeepUs As String)
    If Not KeepUs Like "*Units*" Then Me.Units_List.RowSource = ""
    If Not KeepUs Like "*Lease*" Then Me.LeaseList.RowSource = ""
    If Not KeepUs Like "*Schedule*" Then Me.ScheduleList.RowSource = ""
    If Not KeepUs Like "*Group*" Then Me.GroupList.RowSource = ""
    If Not KeepUs Like "*UnitDesc*" Then Me.DescList.RowSource = ""
End Function

Public Function ToggleSort(toggleField As String)
    If toggleField <> "AssetID" Then Me.ToggleSortAssetID = 0
    If toggleField <> "UnitVIN" Then Me.ToggleSortUnitVIN = 0
    If toggleField <> "UnitNum" Then Me.ToggleSortUnitNum = 0
    If toggleField <> "UnitDesc" Then Me.ToggleSortUnitDesc = 0
    If toggleField = "ClearAll" Then Exit Function
    SQLz = SplitString(Units_List.RowSource, " Order By", 0) & " Order By " & toggleField & " "
    If Me("ToggleSort" & toggleField) = 0 Then SQLz = SQLz & " Desc"
    Me.Units_List.RowSource = SQLz
End Function
Private Sub ToggleSortAssetID_Click()
   Call ToggleSort("AssetID")
End Sub

Private Sub ToggleSortUnitDesc_Click()
   Call ToggleSort("UnitDesc")
End Sub

Private Sub ToggleSortUnitNum_Click()
   Call ToggleSort("UnitNum")
End Sub

Private Sub ToggleSortUnitVIN_Click()
   Call ToggleSort("UnitVIN")
End Sub

Private Sub Units_List_Click()
    If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.Units_List.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Pick_Units_WorkTable(UnitID) SELECT " & Me.Units_List.Column(0) & ";")
        DoCmd.SetWarnings True
        Call ReselectPickedUnits
    End If
    Call GetUnitCounts
    Me.GroupList = DLookup("SGrpId", "vw_SixKeys", "UnitID=" & Me.Units_List.Column(0) & "")
End Sub


Public Sub ReselectPickedUnits()
    Dim i, x As Integer
    Dim pUnits As String
    Dim RST As DAO.Recordset
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
Public Sub ListboxSelectString(lstListbox As ListBox, strToSelect As String, Optional SelectDeselect As Boolean = True)
    Dim i As Integer
    For i = 0 To lstListbox.ListCount - 1
        If lstListbox.ItemData(i) = strToSelect Then
            lstListbox.Selected(i) = SelectDeselect
        End If
    Next i
End Sub
Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount]) > 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
    Me.GetAll.caption = ChrW(9660) & "All " & ([Units_List].ListCount) & ""    '& UnitX
    If ([Units_List].[ListCount]) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
    Me.GetSelectedUnits.caption = ChrW(9660) & "Selected (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ
    If [Units_List].[ItemsSelected].[Count] > 0 Then Me.GetSelectedUnits.Enabled = True Else Me.GetSelectedUnits.Enabled = False
End Sub

Private Sub COFLeaseReport_Click()
    If Me.GrpID & "test" = "test" Then
        MsgBox "You must Select a Group before you can Request a Certificate of Lease", vbCritical, "No Group Selected"
        Exit Sub
    End If
    DoCmd.close acReport, "CertOfLease_Mass"
    DoCmd.OpenReport "CertOfLease_Mass", acViewPreview
End Sub
Private Sub CofA_Report_Click()
'    On Error GoTo ExitMe
    If Me.GrpID & "test" = "test" Then
        MsgBox "You must Select a Group before you can Request a CofA", vbCritical, "No Group Selected"
        Exit Sub
    End If
    If Me.UnitCount < 1 Then
        MsgBox "You must Select at lease One Unit before you can request a CofA", vbCritical, "No Units Selected"
        Exit Sub
    End If
    DoCmd.close acForm, "CofASignerPopUp"
    DoCmd.OpenForm "CofASignerPopUp", acNormal, , , , , Me.GrpID
End Sub
Private Sub SUF_Mass_Click()
    If DCount("UnitId", "Bulk_Update_WorkTable") < 1 Then
        MsgBox "You must Select Units before you can request an SUF", vbCritical, "No Units Selected"
        Exit Sub
    End If
    Dim Queryx As String
    Queryx = "vw_SUFReportQuery"
    ReportX = "SUFReport"
    If DCount("ClientID", "SUFConditions", "ClientID=" & Me.ClientID) > 0 Then ReportX = "SUFwConditions"
    If IsAutohauler("SchId", Me.ScheduleList) > 0 Then
        Queryx = "vw_SUFReportAutoHauler"
        ReportX = "SUFReport_AutoHauler"
    End If
    If Me.ClientList.Column(1) Like "*Maines*" Then ReportX = "SUFReport_Maines"
    If Me.ClientID = 4971 Then ReportX = "SUFMass_Mclane"
    SQLz = "SELECT * "
    SQLz = SQLz & "FROM " & Queryx & " "
    SQLz = SQLz & "INNER JOIN Bulk_Update_WorkTable "
    SQLz = SQLz & "ON " & Queryx & ".UnitID = Bulk_Update_WorkTable.UnitID "
    SQLz = SQLz & "ORDER BY UnitVin;"
    'ReportX = "SUFMass"
    If ReportX = "" Then Exit Sub
MakeReport:
    DoCmd.OpenReport ReportX, acViewPreview, , , , SQLz
ExitMe:
End Sub
Private Sub CloseMe_Click()
    DoCmd.SetWarnings False
        If Me.LeaseList.Column(2) = "FA" Then
        Call CheckForReasons
        Call RePairExchangeIT
        End If
    DoCmd.SetWarnings True
    DoCmd.close acForm, Me.Form.Name
End Sub
Private Sub Form_Close()
    On Error Resume Next
    Call SetOrderBy("[Bulk_Update_SelectedUnits].[UnitID]")
    On Error GoTo FixMe
    DoCmd.OpenForm " FAMainMenu"
FixMe:
End Sub

Public Sub GetAll_Click()
    For i = 0 To Units_List.ListCount - 1
        Units_List.Selected(i) = True
    Next i
    Call GetSelectedUnits_Click
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub GetSelectedUnits_Click()
    Call SetUnits
End Sub
Public Function SetOrderBy(fOrderBy As String)

If Not fOrderBy Like "*Look*" Then
    If Not fOrderBy Like "*_SYND*" Then
    Forms!BulkChange_Form.BulkChange_CoaMso.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_CoaMso.Form.OrderByOn = True

    Forms!BulkChange_Form.BulkChange_Desc.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_Desc.Form.OrderByOn = True

    Forms!BulkChange_Form.BulkChange_Desc2.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_Desc2.Form.OrderByOn = True

    Forms!BulkChange_Form.BulkChange_Rents.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_Rents.Form.OrderByOn = True

    Forms!BulkChange_Form.BulkChange_Vendor.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_Vendor.Form.OrderByOn = True

    Forms!BulkChange_Form.BulkChange_Syndication.Form.OrderBy = Replace(fOrderBy, "Bulk_Update_SelectedUnits", "Bulk_Update_SelectedUnits_SYND")
    Else
    Forms!BulkChange_Form.BulkChange_Syndication.Form.OrderByOn = True
    End If
    If fOrderBy = "[Bulk_Update_SelectedUnits].[UnitID]" _
       Or fOrderBy = "[Bulk_Update_SelectedUnits].[UnitVIN]" _
       Or fOrderBy = "[Bulk_Update_SelectedUnits].[UnitStatus]" Then
        Forms!BulkChange_Form.MassMod_Titles.Form.OrderBy = Replace(Replace(Replace(fOrderBy, "Bulk_Update_SelectedUnits", "UnitTitleLog_Work"), "UnitId", "AssetID"), "UnitVIN", "VIN")
        Forms!BulkChange_Form.MassMod_Titles.Form.OrderByOn = True
    End If
End If
If Not fOrderBy Like "*_SYND*" Then
    Forms!BulkChange_Form.BulkChange_Locations.Form.OrderBy = fOrderBy
    Forms!BulkChange_Form.BulkChange_Locations.Form.OrderByOn = True
    On Error GoTo 0
End If
End Function
Public Function SetUnits(Optional fFilter As String)
    If fFilter Like "*LState*" Then Exit Function
    Dim ItemIndex As Variant
    Dim UnitIDx As String
    Call DelWorkTable
    For Each ItemIndex In Me.Units_List.ItemsSelected
        UnitIDx = UnitIDx & Me.Units_List.ItemData(ItemIndex) & ","
    Next
    If Right(UnitIDx, 1) = "," Then UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
    ' Make Mass-Mod Work table.
    SQLx = "INSERT INTO Bulk_Update_WorkTable (UnitID, UnitDesc, ChildPartCost, DrawDownID) " & _
        "Select UnitID, [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitModel] & ' ' & [UnitSubType] & ' ' & [UnitType] " & _
        ",nz(CPC,0), Units.DrawDownID From Units " & _
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
    SQLx = SQLx & ", IIf([EntryTypeId]=1,'IN',IIf([EntryTypeID]=2,'Out','-')) AS Dir, '" & FaUserName & "' AS [User]"
    SQLx = SQLx & ", nz(H.EntryReason & ' (' & H.UnitTitleLogDate &')','N/A')"
    SQLx = SQLx & ",EntryTypeId, [Status] "
    SQLx = SQLx & "FROM (vw_SixKeys AS S INNER JOIN Units AS U ON S.AssetID = U.UnitID) "
    SQLx = SQLx & "LEFT JOIN UnitTitleHistoryLastEntry AS H ON U.UnitID = H.AssetID "
    SQLx = SQLx & "WHERE (((S.AssetID) In (Select UnitID From Bulk_Update_WorkTable)));"
    DoCmd.RunSQL (SQLx)

    SQLx = ""
    DoCmd.SetWarnings True

    Me.ClientID = Me.ClientList.Column(0)
    Call ReQueryAll
    On Error GoTo fixme1
    Me.UnitCount.Visible = True
    Me.UnitCount = Me.BulkChange_Desc.Form!UnitCount
    Exit Function
fixme1:
End Function
Public Sub ReQueryAll()
    Me!BulkChange_Desc.Form.ClearBoxes
    Me!BulkChange_Desc2.Form.ClearBoxes
    Me!BulkChange_Rents.Form.ClearBoxes
    Me!BulkChange_Vendor.Form.ClearBoxes
    Me!BulkChange_Locations.Form.ClearBoxes
    Me!BulkChange_CoaMso.Form.ClearBoxes
    Me!BulkChange_ChildParts.Form.ClearBoxes
    Me!BulkChange_ChildPartControl.Form.ClearBoxes
    Me!MassMod_Titles.Form.ClearBoxes
    Me!BulkChange_Syndication.Form.ClearBoxes
    Me.BulkChange_Desc.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_Desc2.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_Rents.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_Vendor.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_Locations.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_CoaMso.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_ChildParts.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me.BulkChange_ChildPartControl.Form.RecordSource = "Bulk_Update_SelectedUnits"
    Me!BulkChange_Syndication.Form.RecordSource = "Bulk_Update_SelectedUnits_SYND"
    Me!BulkChange_Syndication.Form.Requery
    Me.BulkChange_ChildParts.Form.ChildList.Requery
    Me.MassMod_Titles.Form.Requery
    Me.BulkChange_ChildPartControl.Form.ChildList.Requery
    Me.BulkChange_ChildPartControl.Form.SisterParts.Requery
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
Private Sub updateGetall()
    Me.GetAll.caption = ChrW(9660) & "All " & [Units_List].ListCount
End Sub
Private Sub SearchAssetID_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchUnitVin_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchUnitNum_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchUnitDesc_LostFocus()
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName AS Client, Count(AssetID) AS Units, ClientGroupId " & _
           "From vw_SixKeys WHERE ClientShortName like '*" & SearchClient.Text & "*' or CompanyName like '*" & SearchClient.Text & "*' " & _
           "GROUP BY clientID, ClientShortName, ClientGroupId " & _
           "ORDER BY vw_SixKeys.ClientShortName;"
    Me.ClientList.RowSource = SQLz
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Public Function SearchUnits_List(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetID = ""
    If sField <> "UnitVIN" Then Me.SearchUnitVin = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum = ""
    If sField <> "UnitDesc" Then Me.SearchUnitDesc = ""
    If sField = "ClearAll" Then Exit Function
    If Len(Me("Search" & sField).Text) >= 1 And Len(Me("Search" & sField).Text) < 3 Then Exit Function
    SQLz = "SELECT AssetID, UnitVin, UnitNum, UnitDesc, SGrpID " & _
           "From Bulk_Units_FAID_SCH " & _
           "WHERE SchID =" & Me.ScheduleList.Column(0) & " AND " & sField & " Like '*" & Me("Search" & sField).Text & "*' "
            If Me.GrpID <> "" Then SQLz = SQLz & " AND SGrpID =" & Me.GrpID & " "
            SQLz = SQLz & " Order By " & sField & ";"
    Me.Units_List.RowSource = SQLz
    Call updateGetall
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Function

Private Sub SearchAssetId_Change()
    Call SearchUnits_List("AssetID")
End Sub
Private Sub SearchUnitVin_Change()
    Call SearchUnits_List("UnitVIN")
End Sub
Private Sub SearchUnitNum_Change()
    Call SearchUnits_List("UnitNum")
End Sub
Private Sub SearchUnitDesc_Change()
    Call SearchUnits_List("UnitDesc")
End Sub

Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub UpdatePortal_Click()
'On Error GoTo fixme
    Dim RST As DAO.Recordset
    SQLz = "Select * From UnitWorkTable;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    If RST.RecordCount < 1 Then
FixMe:
        MsgBox "There are no records to update", , "No Data"
        GoTo Cleanup
    Else
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "SQLSP_UpdateUnitsDateStatus"
        DoCmd.SetWarnings True
        MsgBox "Data Sent to Portal", , "Data Sent :)"
    End If
Cleanup:
    RST.close
    Set RST = Nothing
End Sub
