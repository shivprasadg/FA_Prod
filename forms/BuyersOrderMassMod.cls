Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim RST        As DAO.Recordset
Dim RSTunit        As DAO.Recordset
Public Sub ClearBOSearchxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub ClearFilters_Click()
  '  Me.ReMarketingMassMod_subform1.Form.FilterOn = False
End Sub
Private Sub HideSubs()
    Me.ReMarketingMassMod_sub_SUFParked.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_BuyOutView.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_FMVPenalty.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_SoldInvoice.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_SUFParked.Visible = False
    Me.ReMarketingMassMod_sub_BuyOutView.Visible = False
    Me.ReMarketingMassMod_sub_FMVPenalty.Visible = False
    Me.ReMarketingMassMod_sub_SoldInvoice.Visible = False
    Me.ReMarketingMassMod_sub_Inspection2.Visible = False
    Me.RemarketingMassMod_sub_Titles.Visible = False

End Sub
Private Sub ShowSubs()

   SQLz = "SELECT"
   SQLz = SQLz & " R.*"
   SQLz = SQLz & ", U.UnitID"
   SQLz = SQLz & ", U.UnitUnitNum AS UnitNum"
   SQLz = SQLz & ", U.UnitVIN"
   SQLz = SQLz & ", U.UnitStatus"
   SQLz = SQLz & ", U.UnitOffLeaseDt"
   SQLz = SQLz & ", U.LocationParkedID"
   SQLz = SQLz & ", U.ClientGroupID"
   SQLz = SQLz & ", SG.AssigneeBankID"
   SQLz = SQLz & ", S.SchID"
   SQLz = SQLz & ", S.SchNo"
   SQLz = SQLz & ", BO.Last"
   SQLz = SQLz & " FROM  ((((RMKT As R"
   SQLz = SQLz & " INNER JOIN BuyersOrder_MMod_Work AS BO ON R.UnitRef = BO.UnitID)"
   SQLz = SQLz & " INNER JOIN Units AS U     ON R.UnitRef = U.UnitID)"
   SQLz = SQLz & " INNER JOIN SchGrp AS SG   ON U.GroupREF  = SG.SGrpID)"
   SQLz = SQLz & " INNER JOIN Schedule AS S  ON SG.SGrpSchID = S.SchID);"
  '' SQLz = SQLz & " LEFT JOIN RMKTUnitTitleLog_Work AS LT ON LT.AssetID = U.UnitID;"
 '' SQLz = SQLz & " WHERE (((Units.UnitStatus) <> 'B'));"

    Me.ReMarketingMassMod_sub_SUFParked.Visible = True
    Me!ReMarketingMassMod_sub_SUFParked.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_SUFParked.Form.RecordSource = SQLz

    Me.ReMarketingMassMod_sub_BuyOutView.Visible = True
    Me!ReMarketingMassMod_sub_BuyOutView.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_BuyOutView.Form.RecordSource = SQLz

    Me.ReMarketingMassMod_sub_FMVPenalty.Visible = True
    Me!ReMarketingMassMod_sub_FMVPenalty.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_FMVPenalty.Form.RecordSource = SQLz

    Me.ReMarketingMassMod_sub_SoldInvoice.Visible = True
    Me.ReMarketingMassMod_sub_SoldInvoice.Form.RecordSource = SQLz

    SQLz = " SELECT UnitInspections.* "
    SQLz = SQLz & " FROM UnitInspections"
    SQLz = SQLz & " INNER JOIN UnitInsp_WorkTable "
    SQLz = SQLz & " ON UnitInspections.InspnID = UnitInsp_WorkTable.InspnID "
    SQLz = SQLz & " ORDER BY UnitInspections.AssetID, UnitInspections.InspNum DESC;"
    Me.ReMarketingMassMod_sub_Inspection2.Visible = True
    Me.ReMarketingMassMod_sub_Inspection2.Form.RecordSource = SQLz

    Me.RemarketingMassMod_sub_Titles.Visible = True
    Me.RemarketingMassMod_sub_Titles.Form.RecordSource = "RMKTUnitTitleLog_Work"
    Me.RemarketingMassMod_sub_Titles.Form.Requery

End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.BuyNo = "Please select a Buyers Order..."

     SQLz = "SELECT InvoiceRequestID as BO"
     SQLz = SQLz & " ,BuyerShortName AS Buyer"
     SQLz = SQLz & " ,InvoicePaidDate AS Paid"
     SQLz = SQLz & " ,SP"
     SQLz = SQLz & " ,Count(UnitID) as Units"
     SQLz = SQLz & " FROM vw_BuyersOrder"
     SQLz = SQLz & " WHERE WON = True "
     SQLz = SQLz & " GROUP BY InvoiceRequestID, BuyerShortName,InvoicePaidDate,SP"
     SQLz = SQLz & " ORDER BY InvoiceRequestID DESC;"

    Me.BuyersOrderList.RowSource = SQLz
    Call ClearBOSearch("all")
    Call ClearUnitSearch("all")
    Call DelWorkTable
    Me.Units_List.RowSource = ""
    Me.UnitCount.Visible = False
    Call HideSubs
End Sub

Public Sub GetAll_Click()
   Call BuyersOrderList_DblClick(1)
End Sub

Private Sub BuyersOrderList_DblClick(Cancel As Integer)
    Call DelWorkTable
    ' Make Unit Work Table.
    DoCmd.SetWarnings False
    SQLz = "INSERT INTO BuyersOrder_MMod_Work (UnitID, BuyerID, LastEntry, Last) "
    SQLz = SQLz & "SELECT UnitID, BuyerID, EntryReason, IIf([EntryTypeId]=1,'IN',IIf([EntryTypeID]=2,'Out','-')) AS Last "
    SQLz = SQLz & "FROM vw_BuyersOrderUnits "
    SQLz = SQLz & "LEFT JOIN UnitTitleHistoryLastEntry AS H ON vw_BuyersOrderUnits.UnitID = H.AssetID "
    SQLz = SQLz & "WHERE BuyersOrderID=" & Me.BuyersOrderList & " "
    DoCmd.RunSQL (SQLz)
    ' Make Unit Insp Work Table.
    SQLz = "INSERT INTO UnitInsp_WorkTable (InspnID) SELECT Distinct InspnID From UnitInspections Where AssetID in (SELECT UnitID FROM BuyersOrder_MMod_Work)"
    DoCmd.RunSQL (SQLz)

    ' Make Unit Title Work Table.
    SQLx = ""
    SQLx = "INSERT INTO RMKTUnitTitleLog_Work ( AssetID, VIN, TitleNumber, TitleSt, [Desc], LastDirection, UserID, LastEntry, LastEntryTypeId, UnitStatus) "
    SQLx = SQLx & "SELECT S.AssetID, S.VIN, H.TitleNumber, S.TitledState, S.UnitDescShortVer"
    SQLx = SQLx & ", IIf([EntryTypeId]=1,'IN',IIf([EntryTypeID]=2,'Out','-')) AS Dir, '" & FAUserName & "' AS [User]"
    SQLx = SQLx & ", nz(H.EntryReason & ' (' & H.UnitTitleLogDate &')','N/A')"
    SQLx = SQLx & ", EntryTypeId, [Status] "
    SQLx = SQLx & "FROM (vw_SixKeys AS S INNER JOIN Units AS U ON S.AssetID = U.UnitID) "
    SQLx = SQLx & "LEFT JOIN UnitTitleHistoryLastEntry AS H ON U.UnitID = H.AssetID "
    SQLx = SQLx & "WHERE (((S.AssetID) In (Select UnitID From BuyersOrder_MMod_Work)));"
    DoCmd.RunSQL (SQLx)
    Call ShowSubs
    Call ClearBOSearch("all")
    Call ClearUnitSearch("all")
    Call GetUnitCounts
    Me.UnitCount.Visible = True
    DoCmd.SetWarnings True
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From BuyersOrder_MMod_Work;")
    DoCmd.RunSQL ("DELETE * From UnitInsp_WorkTable;")
    DoCmd.RunSQL ("DELETE * From RMKTUnitTitleLog_Work;")
    DoCmd.SetWarnings True
End Sub

Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount] - 1) = 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
    Me.GetAll.caption = ChrW(9660) & "All " & ([Units_List].ListCount - 1) & " Listed " & UnitX
    If ([Units_List].[ListCount] - 1) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
End Sub
Public Sub BuyersOrderList_Click()
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "From vw_BuyersOrder "
    SQLz = SQLz & "Where InvoiceRequestID = " & Me.BuyersOrderList & " "
    SQLz = SQLz & "Order BY InvoiceRequestID DESC; "
    Me.Units_List.RowSource = SQLz
    Call HideSubs
    Call GetUnitCounts
    Call ClearBOSearch("all")
    Call ClearUnitSearch("all")
    Me.BuyNo = "Ord: " & Me.BuyersOrderList.Column(0) & " Buyer: " & Me.BuyersOrderList.Column(1) & " "
End Sub


Private Sub MakeBTF_Click()
    If Me.BuyersOrderList & "test" = "test" Then
          MsgBox "You haven't Selected a Buyers Order" & vbNewLine & "(Wake up!?)", vbCritical, "No BO, No BTF"
          Exit Sub
    End If
    Dim AskDate As Date
On Error GoTo FixMe
AskDate:
        AskDate = InputBox("Enter the Projected BuyOut DATE for Buyers Order: " & Me.BuyersOrderList & " | " & Me.BuyersOrderList.Column(1) & vbNewLine & _
                         "          Please enter that date, below." & vbNewLine & _
                         "          Today's date is the default." & vbNewLine & vbNewLine & _
                         "          ( You MUST use a date )" _
                , "Projected BuyOut Date:", Date)
        If Not IsDate(AskDate) Then GoTo AskDate
    Call SendBuyersOrderToBTF(AskDate, Me.BuyersOrderList)
FixMe:
Exit Sub
On Error GoTo 0
End Sub

Private Sub MaketitleRequest_Click()
    If DCount("UnitID", "BuyersOrder_MMod_Work") = 0 Then
        MsgBox "You haven't Selected a Buyers Order" & vbNewLine & "(Are you daff?)", vbCritical, "No units, No service"
        Exit Sub
    End If
    SQLz = "Select AssigneeBankID, Count(UnitID) as Units From BuyersOrderMassMod1 Group By AssigneeBankID"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveFirst
        Dim i As Integer
        Dim u As Integer
        For i = 1 To RST.RecordCount
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Delete From Pick_Units_WorkTable;")
            DoCmd.RunSQL ("Insert into Pick_Units_WorkTable (UnitID) Select UnitID From BuyersOrderMassMod1 WHERE AssigneeBankID=" & RST!AssigneeBankId & ";")
            DoCmd.SetWarnings True
            Call MakeTitleRequestLetter(RST!AssigneeBankId, 1, "", "TitleReleaseRequest_" & i)
            RST.MoveNext
        Next i
    'Shell "explorer.exe """ & "C:\FAAS\" & FAUserName & "\temp\" & """, vbNormalFocus"
    End If
End Sub
Private Sub MakeSingleTitleRequest_Click()
    If DCount("UnitID", "BuyersOrder_MMod_Work") = 0 Or Nz(Me.UnitID, 0) = 0 Then
        MsgBox "You haven't Selected a Buyers Order or a Unit" & vbNewLine & "(Are you daff?)", vbCritical, "No units, No service"
        Exit Sub
    End If
    SQLz = "Select AssigneeBankID, UnitID From BuyersOrderMassMod1 Where UnitID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount = 0 Then
        MsgBox "There is no title held for AssetID: " & Me.UnitID & vbNewLine & "(Please try another unit", vbCritical, "No units, No service"
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Pick_Units_WorkTable;")
    DoCmd.RunSQL ("Insert into Pick_Units_WorkTable (UnitID) Select " & RST!UnitID & ";")
    DoCmd.SetWarnings True
    Call MakeTitleRequestLetter(RST!AssigneeBankId, 1, "", "TitleReleaseRequest_single")
    'Shell "explorer.exe """ & "C:\FAAS\" & FAUserName & "\temp" & """, vbNormalFocus"
End Sub
Private Sub RMKTTabControl_Change()
    Call SetSubData(Me.RMKTTabControl.Value)
End Sub
Public Function SetSubData(TabID As Integer)
    If TabID = 5 Then    ' Titles Tab
       Me.RemarketingMassMod_sub_Titles.Form.Requery
    End If
End Function

Private Sub ClearClientSearch_Click()
    Call Form_Load
End Sub
Private Sub SearchBO_Change()
    Call ClearBOSearch("SearchBO")
    If Len(Me.SearchBO.Text) < 3 Then Exit Sub
    SQLz = "SELECT InvoiceRequestID as BO"
    SQLz = SQLz & " ,BuyerShortName AS Buyer"
    SQLz = SQLz & " ,InvoicePaidDate AS Paid"
    SQLz = SQLz & " ,SP"
    SQLz = SQLz & " ,Count(UnitID) as Units"
    SQLz = SQLz & " FROM vw_BuyersOrder"
    SQLz = SQLz & " WHERE InvoiceRequestID Like '*" & Me.SearchBO.Text & "*' "
    SQLz = SQLz & " GROUP BY InvoiceRequestID, BuyerShortName,InvoicePaidDate,SP"
    SQLz = SQLz & " ORDER BY InvoiceRequestID DESC;"
    Me.BuyersOrderList.RowSource = SQLz
End Sub

Private Sub SearchBuyer_Change()
    Call ClearBOSearch("SearchBuyer")
    If Len(Me.SearchBuyer.Text) < 3 Then Exit Sub
    SQLz = "SELECT InvoiceRequestID as BO"
    SQLz = SQLz & " ,BuyerShortName AS Buyer"
    SQLz = SQLz & " ,InvoicePaidDate AS Paid"
    SQLz = SQLz & " ,SP"
    SQLz = SQLz & " ,Count(UnitID) as Units"
    SQLz = SQLz & " FROM vw_BuyersOrder"
    SQLz = SQLz & " WHERE BuyerShortName Like '*" & Me.SearchBuyer.Text & "*' "
    SQLz = SQLz & " GROUP BY InvoiceRequestID, BuyerShortName,InvoicePaidDate,SP"
    SQLz = SQLz & " ORDER BY InvoiceRequestID DESC;"
    Me.BuyersOrderList.RowSource = SQLz
End Sub

Private Sub SearchPaid_AfterUpdate()
    Call ClearBOSearch("SearchPaid")
    If IsDate(Me.SearchPaid) Then
        SQLz = "SELECT InvoiceRequestID as BO"
        SQLz = SQLz & " ,BuyerShortName AS Buyer"
        SQLz = SQLz & " ,InvoicePaidDate AS Paid"
        SQLz = SQLz & " ,SP"
        SQLz = SQLz & " ,Count(UnitID) as Units"
        SQLz = SQLz & " FROM vw_BuyersOrder"
        SQLz = SQLz & " WHERE InvoicePaidDate = #" & CDate(Me.SearchPaid) & "# "
        SQLz = SQLz & " GROUP BY InvoiceRequestID, BuyerShortName,InvoicePaidDate,SP"
        SQLz = SQLz & " ORDER BY InvoiceRequestID DESC;"
        Me.BuyersOrderList.RowSource = SQLz
    Else
        Me.SearchPaid = ""
    End If
End Sub

Private Sub SearchSP_Change()
    Call ClearBOSearch("SearchSP")
    If Len(Me.SearchSP.Text) < 2 Then Exit Sub
    SQLz = "SELECT InvoiceRequestID as BO"
    SQLz = SQLz & " ,BuyerShortName AS Buyer"
    SQLz = SQLz & " ,InvoicePaidDate AS Paid"
    SQLz = SQLz & " ,SP"
    SQLz = SQLz & " ,Count(UnitID) as Units"
    SQLz = SQLz & " FROM vw_BuyersOrder"
    SQLz = SQLz & " WHERE SP Like '*" & Me.SearchSP.Text & "*' "
    SQLz = SQLz & " GROUP BY InvoiceRequestID, BuyerShortName,InvoicePaidDate,SP"
    SQLz = SQLz & " ORDER BY InvoiceRequestID DESC;"
    Me.BuyersOrderList.RowSource = SQLz
End Sub
Function ClearBOSearch(Excludeme As String)
    If Excludeme <> Me.SearchBO.Name Then Me.SearchBO = ""
    If Excludeme <> Me.SearchBuyer.Name Then Me.SearchBuyer = ""
    If Excludeme <> Me.SearchPaid.Name Then Me.SearchPaid = ""
    If Excludeme <> Me.SearchSP.Name Then Me.SearchSP = ""
    Me.UnitID = Null
End Function
Private Sub SearchAsset_Change()
    Call ClearUnitSearch("SearchAsset")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE UnitID LIKE '*" & Me.SearchAsset.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitID; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub

Private Sub SearchUnit_Change()
    Call ClearUnitSearch("SearchUnit")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE UnitNum LIKE '*" & Me.SearchUnit.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitNum; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchVIN_Change()
    Call ClearUnitSearch("SearchVin")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE VIN LIKE '*" & Me.SearchVin.Text & "*'"
    SQLz = SQLz & "ORDER BY VIN; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Call ClearUnitSearch("SearchStatus")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE [Status] LIKE '*" & Me.SearchStatus.Text & "*'"
    SQLz = SQLz & "ORDER BY VIN; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchMake_Change()
    Call ClearUnitSearch("SearchMake")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE Make LIKE '*" & Me.SearchMake.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitID; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchType_Change()
    Call ClearUnitSearch("SearchType")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE SubType LIKE '*" & Me.SearchType.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitID; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Function ClearUnitSearch(Excludeme As String)
    If Excludeme <> Me.SearchAsset.Name Then Me.SearchAsset = ""
    If Excludeme <> Me.SearchUnit.Name Then Me.SearchUnit = ""
    If Excludeme <> Me.SearchVin.Name Then Me.SearchVin = ""
    If Excludeme <> Me.SearchStatus.Name Then Me.SearchStatus = ""
    If Excludeme <> Me.SearchMake.Name Then Me.SearchMake = ""
    If Excludeme <> Me.SearchType.Name Then Me.SearchType = ""
    Me.UnitID = Null
End Function
Private Sub ClearLabel_Click()
    Call ClearUnitSearch("all")
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub

Private Sub Units_List_Click()
    Me.UnitID = Me.Units_List.Column(1)
End Sub

Public Sub Units_List_DblClick(Cancel As Integer)
    Me.BuyersOrderList = Me.Units_List.Column(0)
    Call ClearLabel_Click
    Call BuyersOrderList_Click
End Sub
