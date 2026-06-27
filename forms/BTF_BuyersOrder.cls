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
    Me.Units_List.RowSource = ""
End Sub



Public Sub BuyersOrderList_Click()
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
    SQLz = SQLz & "From vw_BuyersOrder "
    SQLz = SQLz & "Where InvoiceRequestID = " & Me.BuyersOrderList & " "
    SQLz = SQLz & "Order BY VIN; "
    Me.Units_List.RowSource = SQLz
    Call ClearBOSearch("all")
    Call ClearUnitSearch("all")
    Me.BuyNo = "Ord: " & Me.BuyersOrderList.Column(0) & " Buyer: " & Me.BuyersOrderList.Column(1) & " "
End Sub

Private Sub MakeBOBTF_Click()
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
Private Sub MakeUnitBTF_Click()
    If Me.Units_List.Column(1) & "test" = "test" Then
          MsgBox "You haven't Selected a Unit" & vbNewLine & "(Wake up!?)", vbCritical, "No Unit, No BTF"
          Exit Sub
    End If
    Dim AskDate As Date
On Error GoTo FixMe
AskDate:
        AskDate = InputBox("Enter the Projected BuyOut DATE for AssetID: " & Me.Units_List.Column(1) & " | " & Me.Units_List.Column(9) & vbNewLine & _
                         "          Please enter that date, below." & vbNewLine & _
                         "          Today's date is the default." & vbNewLine & vbNewLine & _
                         "          ( You MUST use a date )" _
                , "Projected BuyOut Date:", Date)
        If Not IsDate(AskDate) Then GoTo AskDate
    Call RunBTF(AskDate, CStr(Me.Units_List.Column(1)))
FixMe:
Exit Sub
On Error GoTo 0
End Sub

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
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE UnitID LIKE '*" & Me.SearchAsset.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitID; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub

Private Sub SearchUnit_Change()
    Call ClearUnitSearch("SearchUnit")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE UnitNum LIKE '*" & Me.SearchUnit.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitNum; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchVIN_Change()
    Call ClearUnitSearch("SearchVin")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc"
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE VIN LIKE '*" & Me.SearchVin.Text & "*'"
    SQLz = SQLz & "ORDER BY VIN; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Call ClearUnitSearch("SearchStatus")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE [Status] LIKE '*" & Me.SearchStatus.Text & "*'"
    SQLz = SQLz & "ORDER BY VIN; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchMake_Change()
    Call ClearUnitSearch("SearchMake")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
    SQLz = SQLz & "FROM vw_BuyersOrder "
    SQLz = SQLz & "WHERE Make LIKE '*" & Me.SearchMake.Text & "*'"
    SQLz = SQLz & "ORDER BY UnitID; "
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchType_Change()
    Call ClearUnitSearch("SearchType")
    SQLz = "SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MYear, DescriptionMed as uDesc "
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
