Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, Committ, SetString, BoUnits As String
Dim BI As Long

Public Sub AssignedUnitsList_Click()
End Sub

Private Sub AssignedUnitsList_DblClick(Cancel As Integer)
    If IsNull(AssignedUnitsList.Column(4)) Then Exit Sub
    Call OpenUnitsForm(Me.BuyersOrderUnits_Sub.Form.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.BuyersOrderUnits_Sub.Column(4)
    Forms!Unitsform.UnitsTabControl.Value = 6
End Sub

Private Sub ClearUnAssignedUnits_Click()
    Call ClearList(Me.UnAssignedUnitsList.Name)
End Sub

Public Sub ClearList(ListBoxName As String)
    Dim ItemsX As Variant
    With Me(ListBoxName)
        For Each ItemsX In .ItemsSelected
            .Selected(ItemsX) = False
        Next ItemsX
        Me(ListBoxName) = Null
    End With
    Me(ListBoxName).Requery
End Sub



Private Sub Disposition_Click()
    Call UpdateDisposition
End Sub

Private Sub MarkAsSent_Click()
    If Me.InvoiceRequestID.[ItemsSelected].[Count] < 1 Then
        MsgBox "You Must Select an Request", vbInformation, "No Request...you're Silly"
        Exit Sub
    End If
    BI = Me.InvoiceRequestID
    If Not IsNull(DLookup("InvoicePaidDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & BI)) Then
        MsgBox "You cannot modify a Paid Invoice", vbInformation, "Don't Play With Matches"
        Exit Sub
    End If
    If Not IsNull(DLookup("InvoiceSentDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & BI)) Then
        MsgBox "Request Already Sent!", vbInformation, "you'll never learn"
        Exit Sub
    End If
    If DMin("Status", "vw_BuyersOrder", "InvoiceRequestID=" & BI) = "D" Then
        MsgBox "This Buyers Order has units that are 'D'elivered" & vbNewLine & "You Cannot 'Sent' This Buyers Order", vbCritical, "The work is done"
        Exit Sub
    End If
    Committ = MsgBox("You're About To set the Sent Date to: " & Date & " for Buyers Order: " & BI & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Remove Units?")
    If Committ = vbNo Then Exit Sub
    SQLz = "UPDATE RmktInvoiceRequestTbl SET InvoiceSentDate='" & Date & "' WHERE InvoiceRequestID=" & BI & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub


Private Sub MarkAsAccepted_Click()
    If Me.InvoiceRequestID & "test" = "test" Then
        MsgBox "You Must First Select an Buyers Order", vbInformation, "Water Is Wet"
        Exit Sub
    End If
    BI = Me.InvoiceRequestID
    If Not IsNull(DLookup("InvoicePaidDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & BI)) Then
        MsgBox "You cannot modify a Paid Invoice", vbInformation, "Don't Play With Matches"
        Exit Sub
    End If
    If DMin("Status", "vw_BuyersOrder", "InvoiceRequestID=" & BI) = "D" Then
        MsgBox "This Buyers Order has units that are 'D'elivered" & vbNewLine & "You Cannot 'Accept' This Buyers Order", vbCritical, "Already Sold"
        Exit Sub
    End If
    Committ = MsgBox("Setting Buyers Order " & BI & " As 'Accepted by the Buyer' " & _
                     vbNewLine & "Will Invalidate ALL Other Buyers Orders" & _
                     vbNewLine & "With ANY of theses Units." & vbNewLine & _
                     vbNewLine & "Note: These Units witll move to 'PS' Status." & vbNewLine & _
                     vbNewLine & "Do You Agree?", vbInformation + vbYesNo, "You Made A Sale!")
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    'Set Winnig BO Sent Date if null
    If IsNull(DLookup("InvoiceSentDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & BI)) Then
        DoCmd.RunSQL ("UPDATE RmktInvoiceRequestTbl SET InvoiceSentDate='" & Date & "' WHERE InvoiceRequestID=" & BI & ";")
    End If
    'Mark the "Won" units
    DoCmd.RunSQL ("Update RMKT_BuyersOrders Set Active = 1, Won = 1 Where BuyersOrderID=" & BI & "")
    'Mark the "Won" Units as "PS"
    DoCmd.RunSQL ("Update Units Set UnitStatus='PS' Where UnitId in (Select UnitID From vw_BuyersOrder Where InvoiceRequestNum=" & BI & ")")
    'Mark the "Lost" units
    DoCmd.RunSQL ("Update RMKT_BuyersOrders Set Active = 0, Won=0 Where BuyersOrderID <> " & BI & " and UnitID in " & _
                  "(Select UnitID From vw_BuyersOrder as B Where InvoiceRequestNum=" & BI & ")")
    'Mark the "Won" BuyersOrder
    DoCmd.RunSQL ("Update RmktInvoiceRequestTbl Set Accepted = Date(), Inactive = 0 Where InvoiceRequestID=" & BI & "")
    'Mark the "Lost" BuyersOrder(s)
    DoCmd.RunSQL ("Update RmktInvoiceRequestTbl Set Accepted = Null, Inactive = 1 Where InvoiceRequestID in " & _
                  "(Select BuyersOrderID From RMKT_BuyersOrders Where BuyersOrderID <> " & BI & " and UnitID in " & _
                  "(Select UnitID From vw_BuyersOrder as B Where InvoiceRequestNum=" & BI & "))")
    DoCmd.SetWarnings True
    Me.BuyersOrderUnits_Sub.Requery
    Me.UnAssignedUnitsList.Requery
    Me.InvoiceRequestID.Requery
SendEmail:
    DoCmd.OutputTo acOutputReport, "BuyersOrder", acFormatPDF, "\\fa-poseidon\Atlaas\Temp\BuyersOrder_" & BI & ".PDF"
    Call SendBOEmail("Buyers Order Accepted", BI, "\\fa-poseidon\Atlaas\Temp\BuyersOrder_" & BI & ".PDF")
    Call SendBOEmail("Losing BuyersOrders", BI)
End Sub


Private Sub MarkAsDead_Click()
    If Me.InvoiceRequestID & "test" = "test" Then
        MsgBox "You Must First Select an Buyers Order", vbInformation, "Sorry About That"
        Exit Sub
    End If
    BI = Me.InvoiceRequestID
    Committ = MsgBox("Setting Buyers Order " & BI & " As 'Dead' " & _
                     vbNewLine & "Will Invalidate ONLY THIS Buyers Order" & _
                     vbNewLine & "This will not impact any others Buyers Order" & vbNewLine & _
                     vbNewLine & "Note: If these units are not on any other Order," & _
                     vbNewLine & "They Will Move to 'R' Status." & vbNewLine & _
                     vbNewLine & "Do You Agree?", vbInformation + vbYesNo, "What went wrong?")
    If Committ = vbNo Then Exit Sub
End Sub

Private Sub MarkAsDelivered_Click()
    If FAUserAccess < 15 Then
        MsgBox "Sorry you cannot Mark units 'Delivered'", vbCritical, "Sorry, rules..."
        Exit Sub
    End If
    If Me.InvoiceRequestID & "test" = "test" Then
        MsgBox "You Must First Select an Buyers Order", vbInformation, "there ARE rules..."
        Exit Sub
    End If
    BI = Me.InvoiceRequestID
    If DMin("Status", "vw_BuyersOrder", "InvoiceRequestID=" & BI) = "D" Then
        MsgBox "This Buyers Order is Already 'D'elivered", vbCritical, "The work is done"
        Exit Sub
    End If
    If IsNull(DLookup("InvoicePaidDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & BI)) Then
        MsgBox "You Cannot 'Deliver' Not Paid For Units", vbCritical, "You just can't give stuff away"
        Exit Sub
    End If
    Committ = MsgBox("You are setting the units in Buyers Order: " & BI _
                   & vbNewLine & "to 'D' status, confirming that you have received" & vbNewLine & "a SIGNED Delivery Receipt" & vbNewLine & vbNewLine & _
                     "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Job well done!")
    If Committ = vbNo Then Exit Sub

    DoCmd.SetWarnings False
    'Mark the Units as "D"
    DoCmd.RunSQL ("Update Units Set UnitStatus='D' Where UnitId in (Select UnitID From vw_BuyersOrder Where InvoiceRequestNum=" & BI & ")")
    DoCmd.SetWarnings True
    Me.BuyersOrderUnits_Sub.Requery
    Me.UnAssignedUnitsList.Requery
    Me.InvoiceRequestID.Requery
    Call SendBOEmail("Units Delivered", BI)
End Sub

Private Sub PrintBuyersOrder_Click()
    If Me.InvoiceRequestID & "test" = "test" Then
        MsgBox "You Must First Select an Request", vbInformation, "You're click-happy"
        Exit Sub
    End If
    DoCmd.close acReport, "BuyersOrder"
    DoCmd.OpenReport "BuyersOrder", acViewPreview
End Sub

Private Sub Form_Current()
    SQLz = "SELECT RI.InvoiceRequestID AS Num, RI.InvoiceRequestDate AS Requested, Count(B.UnitID) AS Units"
    SQLz = SQLz & ", RI.BuyerID, RI.InvoiceSentDate AS Sent, RI.InvoicePaidDate AS Paid "
    SQLz = SQLz & "FROM RmktInvoiceRequestTbl AS RI "
    SQLz = SQLz & "LEFT JOIN RMKT_BuyersOrders AS B ON RI.InvoiceRequestID = B.BuyersOrderID "
    SQLz = SQLz & "WHERE BuyerId= " & Nz(Me.Parent.BuyerID, 0) & " "
    SQLz = SQLz & "GROUP BY RI.InvoiceRequestID, RI.InvoiceRequestDate, RI.BuyerID, RI.InvoiceSentDate, RI.InvoicePaidDate "
    SQLz = SQLz & "Order by InvoiceRequestID Desc;"
    BI = 0
    Me.InvoiceRequestID.RowSource = SQLz
    Call GetBoUnits
    Me.AssetSearch = ""
    Me.UnitNumSearch = ""
    Me.DescSearch = ""
    Me.VINSearch = ""

End Sub

Public Sub InvoiceRequestID_Click()
    Dim KeepFilter As String
    KeepFilter = ""
    If Me.AssetSearch <> "" Then KeepFilter = "AssetSearch"
    If Me.UnitNumSearch <> "" Then KeepFilter = "UnitNumSearch"
    If Me.DescSearch <> "" Then KeepFilter = "DescSearch"
    If Me.VINSearch <> "" Then KeepFilter = "VINSearch"
    ' If IsNull(Me.TakeDelivery) Then Me.TakeDelivery = DateAdd("M", 1, DateSerial(Year(Date), Month(Date), 1) + 24)
    Call ClearFilters(KeepFilter)
    SQLz = "SELECT * FROM RmktInvoiceRequestTbl WHERE InvoiceRequestID=" & Me.InvoiceRequestID & ";"
    Me.Form.RecordSource = SQLz
    BuyersOrderUnits_Sub.Visible = True
    Me.BuyersOrderUnits_Sub.Requery
    Call ClearList(Me.UnAssignedUnitsList.Name)
End Sub

Private Sub PrevPurchases_Click()
    DoCmd.OpenReport "RbuyerUnits", acViewPreview
End Sub

Private Sub NewBuyersOrder_Click()
    Committ = MsgBox("You're about to Create a NEW Buyers Order, Are you sure?", vbQuestion + vbYesNo + vbCritical, "More Sales!")
    If Committ = vbNo Then Exit Sub
    Me.BuyersOrderUnits_Sub.Visible = False
    DoCmd.GotoRecord , , acNewRec
    Me.BuyerID = Me.Parent.BuyerID
    Me.RequestedBy = FaUserName
    If IsNull(Me.TakeDelivery) Then Me.TakeDelivery = DateAdd("M", 1, DateSerial(Year(Date), Month(Date), 1) + 24)
End Sub
Private Sub SaveUpdateOrder_Click()
    If Me.BuyerID & "test" = "test" Then Me.BuyerID = Me.Parent.BuyerID
    If Me.SalesRep & "test" = "test" Or Me.RequestDate & "test" = "test" Or IsNull(Me.TakeDelivery) Or Me.Disposition & "test" = "test" Then
        MsgBox "Each Buyers Order Requires a:" & vbNewLine & "A Sales Rep" & vbNewLine & "A Request Date" & vbNewLine & "A Take Delivery Date" & vbNewLine & "Disposition" & vbNewLine, , "No Stuff,why?"
        Exit Sub
    End If
    Committ = MsgBox("You're Saving a Buyers Order For: " & _
                     vbNewLine & DLookup("BuyerShortName", "Buyer", "BuyerID=" & Me.Parent.BuyerID) & " " & _
                     vbNewLine & "  Take Delivery of: " & Me.TakeDelivery & _
                     vbNewLine & "  Disposition of: " & Me.Disposition.Column(1) & _
                     vbNewLine & "  Tax Rate of: " & Format(Me.TaxRateToUse, "Percent") & _
                     vbNewLine & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo, "Add A Buyers Order?")
    Me.Form.Requery
    If Committ = vbNo Then Exit Sub
    Call UpdateDisposition
    Me.InvoiceRequestID.Requery
End Sub
Private Sub UpdateDisposition()
       If Me.Disposition & "test" <> "test" And DCount("UnitID", "RMKT_BuyersOrders", "BuyersOrderID=" & Me.InvoiceRequestID & "") > 0 Then
        SQLz = "Update RMKT Set Wholesale=" & Me.Disposition & " " & _
               "WHERE (Wholesale is null AND UnitRef in (Select UnitID From RMKT_BuyersOrders WHERE BuyersOrderID=" & Me.InvoiceRequestID & ")) " & _
               "OR (Wholesale <> " & Me.Disposition & " AND  UnitRef in (Select UnitID From RMKT_BuyersOrders WHERE BuyersOrderID=" & Me.InvoiceRequestID & "));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
End Sub
Private Sub DeleteBuyersOrder_Click()
    If Me.InvoiceRequestID.[ItemsSelected].[Count] < 1 Then
        MsgBox "You Must Select an Invoice Request", vbInformation, "No Invoice Request"
        Exit Sub
    End If
    BI = Me.InvoiceRequestID
    If DCount("UnitID", "vw_BuyersOrder", "InvoiceRequestNum=" & BI) > 0 Then
        MsgBox "Delete aborted!" & vbNewLine & "Invoice Request " & BI & " has " & DCount("Num", "RemarketingRequestInvoiceUnitsAssigned", "Num=" & BI) & " Unit(s)", vbCritical, "Delete Aborted"
        Exit Sub
    End If
    Committ = MsgBox("Deleting Request: '" & BI & "'" & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Delete it?")
    If Committ = vbNo Then Exit Sub
    SQLz = "Delete From RmktInvoiceRequestTbl where InvoiceRequestID=" & BI & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.InvoiceRequestID.Requery
End Sub
Private Sub AddToProposal_Click()
    If Me.InvoiceRequestID.[ItemsSelected].[Count] < 1 Then
        MsgBox "You Must Select an Request", vbInformation, "No Request...you're Silly"
        Exit Sub
    End If
    If Me.UnAssignedUnitsList.[ItemsSelected].[Count] < 1 Then
        MsgBox "You Must Select One or More of the Availible Units", vbInformation, "Missing Information"
        Exit Sub
    End If
    If Not IsNull(DLookup("InvoicePaidDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & Me.InvoiceRequestID)) Then
        MsgBox "You cannot add or remove units from a Paid Invoice", vbInformation, "No Can Do"
        Exit Sub
    End If
    ' If Not IsNull(DLookup("InvoiceSentDate", "RmktInvoiceRequestTbl", "InvoiceRequestID=" & Me.InvoiceRequestID)) Then
    '     MsgBox "You cannot add or remove units from a Sent Invoice", vbInformation, "No Can Do"
    '     Exit Sub
    ' End If
    If Me.SalesRep & "test" = "test" Then
        MsgBox "You Must Select A Sales Rep", , "No Rep"
        Exit Sub
    End If
    Dim ItemIndex As Variant
    Dim teststatus As String
    DoCmd.SetWarnings False
    For Each ItemIndex In Me.UnAssignedUnitsList.ItemsSelected
        teststatus = DLookup("UnitStatus", "Units", "UnitID=" & Me.UnAssignedUnitsList.ItemData(ItemIndex))
        If teststatus = "EW" Then
           Call MsgBox("You're trying to add a Unit with the '" & teststatus & "' Status!" & vbNewLine & "You must close the Work Oder before you can proceed.", vbCritical, "Adding a " & teststatus & " Status to a BO!")
           GoTo ContinueLoop:
        End If
        If teststatus = "EF" Then
           Call MsgBox("You're trying to add a Unit with the '" & teststatus & "' Status!" & vbNewLine & "You must enter a PASSSED Inspn before you can proceed.", vbCritical, "Adding a " & teststatus & " Status to a BO!")
           GoTo ContinueLoop:
        End If
        SQLz = "Insert Into RMKT_BuyersOrders (BuyersOrderID, UnitID, SP, SalePrice ) "
        SQLz = SQLz & "Select " & Me.InvoiceRequestID & ""
        SQLz = SQLz & "," & Me.UnAssignedUnitsList.ItemData(ItemIndex) & ""
        SQLz = SQLz & ",'" & Me.SalesRep.Column(0) & "'"
        SQLz = SQLz & ",Nz(Rmkt.SoldPrice,0) "
        SQLz = SQLz & "FROM Rmkt "
        SQLz = SQLz & "WHERE Rmkt.UnitRef=" & Me.UnAssignedUnitsList.ItemData(ItemIndex) & ";"

        If Nz(DLookup("UnitRef", "RMKT", "UnitREf=" & Me.UnAssignedUnitsList.ItemData(ItemIndex) & ""), 0) = 0 Then
           DoCmd.RunSQL ("Insert into RMKT(UnitREf) Select " & Me.UnAssignedUnitsList.ItemData(ItemIndex) & ";")
        End If
DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
DoCmd.SetWarnings True
ContinueLoop:
    Next
    Call GetBoUnits
    Call UpdateDisposition
    Me.BuyersOrderUnits_Sub.Requery
    Me.UnAssignedUnitsList.Requery
    Me.InvoiceRequestID.Requery
    DoCmd.SetWarnings True
End Sub


Private Sub PrintRequestedInvoice_Click()
    If Me.InvoiceRequestID & "test" = "test" Then
        MsgBox "You Must Select an Invoice Request", vbInformation, "No Request Selected"
        Exit Sub
    End If
    DoCmd.OpenReport "RmktInvoiceRequestRpt", acViewPreview

End Sub

Private Sub SearchRequestedForm_Click()
    DoCmd.close acForm, "RmktInvoiceRequestReverseLookUpForm"
    DoCmd.OpenForm "RmktInvoiceRequestReverseLookUpForm"
End Sub

Private Sub UnAssignedUnitsList_DblClick(Cancel As Integer)
    If IsNull(UnAssignedUnitsList.Column(2)) Then Exit Sub
    Call OpenUnitsForm(Me.UnAssignedUnitsList.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.UnAssignedUnitsList.Column(0)
    Forms!Unitsform.UnitsTabControl.Value = 6
End Sub
Private Sub ClearLabel_Click()
    Call ClearFilters("")
End Sub
Private Function ClearFilters(Optional KeepMe As String)
    If KeepMe = "" Then Me.AssetSearch = ""
    If KeepMe = "" Then KeepMe = "AssetSearch"
    Me(KeepMe).SetFocus
    If KeepMe = "AssetSearch" Then Call AssetSearch_Change
    If KeepMe = "UnitNumSearch" Then Call UnitNumSearch_Change
    If KeepMe = "DescSearch" Then Call UnitNumSearch_Change
    If KeepMe = "VINSearch" Then Call VINSearch_Change
End Function
Public Function GetBoUnits()
    If Nz(Me.InvoiceRequestID, 0) = 0 Then
        BoUnits = "0"
        Exit Function
    End If
    Dim RST    As DAO.Recordset
    Dim i As Integer
    SQLz = "SELECT UnitID From RMKT_BuyersOrders WHERE BuyersOrderID=" & Me.InvoiceRequestID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
    For i = 1 To RST.RecordCount
        If i = 1 Then BoUnits = RST!UnitID & "," Else BoUnits = BoUnits & RST!UnitID & ","
        RST.MoveNext
    Next i
        BoUnits = Left(BoUnits, Len(BoUnits) - 1)
    Else
        BoUnits = "0"
    End If
    RST.close
    Set RST = Nothing
    End Function
Private Sub AssetSearch_Change()
'Me.AssetSearch = ""
    Me.UnitNumSearch = ""
    Me.DescSearch = ""
    Me.VINSearch = ""
    If BoUnits = "" Then Call GetBoUnits

    SQLz = "SELECT RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.SourceCompany, RAU.[Sch|Grp], Count(RAU.AssetID) AS Orders"
    SQLz = SQLz & " FROM vw_RMKT_AvailbleUnits AS RAU"
    SQLz = SQLz & " WHERE (((RAU.[AssetID]) Like '*" & Me.AssetSearch.Text & "*' ))"
    SQLz = SQLz & " GROUP BY RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.[Sch|Grp], RAU.SourceCompany"
    SQLz = SQLz & " HAVING (((RAU.AssetID) Not In (" & BoUnits & ")));"

    Me.UnAssignedUnitsList.RowSource = SQLz
End Sub

Private Sub UnitNumSearch_Change()
    Me.AssetSearch = ""
    'Me.UnitNumSearch = ""
    Me.DescSearch = ""
    Me.VINSearch = ""
    If BoUnits = "" Then Call GetBoUnits

    SQLz = "SELECT RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.SourceCompany, RAU.[Sch|Grp], Count(RAU.AssetID) AS Orders"
    SQLz = SQLz & " FROM vw_RMKT_AvailbleUnits AS RAU"
    SQLz = SQLz & " WHERE (((RAU.[UnitNum]) Like '*" & Me.UnitNumSearch.Text & "*'))"
    SQLz = SQLz & " GROUP BY RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.[Sch|Grp], RAU.SourceCompany"
    SQLz = SQLz & " HAVING (((RAU.AssetID) Not In (" & BoUnits & ")));"

    Me.UnAssignedUnitsList.RowSource = SQLz
End Sub
Private Sub DescSearch_Change()
    Me.AssetSearch = ""
    Me.UnitNumSearch = ""
    'Me.DescSearch = ""
    Me.VINSearch = ""
    If BoUnits = "" Then Call GetBoUnits

    SQLz = "SELECT RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.SourceCompany, RAU.[Sch|Grp], Count(RAU.AssetID) AS Orders"
    SQLz = SQLz & " FROM vw_RMKT_AvailbleUnits AS RAU"
    SQLz = SQLz & " WHERE (((RAU.Description) Like '*" & Me.DescSearch.Text & "*'))"
    SQLz = SQLz & " GROUP BY RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.[Sch|Grp], RAU.SourceCompany"
    SQLz = SQLz & " HAVING (((RAU.AssetID) Not In (" & BoUnits & ")));"

    Me.UnAssignedUnitsList.RowSource = SQLz
End Sub

Private Sub VINSearch_Change()
    Me.AssetSearch = ""
    Me.UnitNumSearch = ""
    Me.DescSearch = ""
    If BoUnits = "" Then Call GetBoUnits

    SQLz = "SELECT RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.SourceCompany, RAU.[Sch|Grp], Count(RAU.AssetID) AS Orders"
    SQLz = SQLz & " FROM vw_RMKT_AvailbleUnits AS RAU"
    SQLz = SQLz & " WHERE (((RAU.VIN) Like '*" & Me.VINSearch.Text & "*'))"
    SQLz = SQLz & " GROUP BY RAU.AssetID, RAU.Unitnum, RAU.VIN, RAU.Description, RAU.Status, RAU.TargetSale, RAU.[Sch|Grp], RAU.SourceCompany"
    SQLz = SQLz & " HAVING (((RAU.AssetID) Not In (" & BoUnits & ")));"

    Me.UnAssignedUnitsList.RowSource = SQLz
End Sub
