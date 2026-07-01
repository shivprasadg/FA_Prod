Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, StringX, PreSQLDrawn, PreSQLOpen, SQLx As String
Dim RST        As DAO.Recordset
Dim OldNum, NewNum As Integer
Dim Commit As String
Dim InvX As String
Dim OpenUnitsX, DrawnUnitsX, OpenPartsX, DrawnPartsX As Integer


Private Sub ApplyEditInvoice_Click()
    If [OpenUnitsList].[ItemsSelected].[Count] < 1 Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From Bulk_Update_WorkTable")
    For Each ItemVar In Me.OpenUnitsList.ItemsSelected
        DoCmd.RunSQL ("INSERT INTO Bulk_Update_WorkTable ( UnitID ) Select " & Me.OpenUnitsList.ItemData(ItemVar) & ";")
    Next ItemVar
    DoCmd.SetWarnings True
    DoCmd.close acForm, "MassMod_InvoiceDraw"
    DoCmd.OpenForm "MassMod_InvoiceDraw"
End Sub


Private Sub ChildAssetSearch_Change()
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
        "WHERE AssetID Like '*" & Me.ChildAssetSearch.Text & "*'" & _
        "AND GroupID=" & Me.GroupID & " " & _
        "Order By AssetID,ChildType,CInvoice;"
    Me.OpenPartsList.RowSource = SQLz
End Sub

Private Sub ChildAssetSearch2_Change()
    ClearSearchBoxes (Me.ActiveControl.Name)
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE AssetID Like '*" & Me.ChildAssetSearch2.Text & "*'" & _
        "AND DrawDownID=" & Me.DrawDownId & " and GroupID=" & Me.GroupID & " " & _
        "Order By AssetID,ChildType,CInvoice;"
    Me.DrawnPartsList.RowSource = SQLz
    Call RunningTotal
End Sub

Private Sub ChildInvoiceSearch_Change()
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
        "WHERE CInvoice Like '*" & Me.ChildInvoiceSearch.Text & "*'" & _
        "AND GroupID=" & Me.GroupID & " " & _
        "Order By CInvoice,AssetID, ChildType;"
    Me.OpenPartsList.RowSource = SQLz
End Sub
Private Sub CInvoiceSearch2_Change()
    ClearSearchBoxes (Me.ActiveControl.Name)
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE CInvoice Like '*" & Me.CInvoiceSearch2.Text & "*'" & _
        "AND DrawDownID=" & Me.DrawDownId & " and GroupID=" & Me.GroupID & " " & _
        "Order By CInvoice,AssetID,ChildType;"
    Me.DrawnPartsList.RowSource = SQLz
    Call RunningTotal
End Sub
Private Sub ChildUnitNumSearch_Change()
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
        "WHERE UnitNum Like '*" & Me.ChildUnitNumSearch.Text & "*'" & _
        "AND GroupID=" & Me.GroupID & " " & _
        "Order By UnitNum,ChildType,CInvoice;"
    Me.OpenPartsList.RowSource = SQLz
End Sub

Private Sub ChildUnitNumSearch2_Change()
    ClearSearchBoxes (Me.ActiveControl.Name)
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE UnitNum Like '*" & Me.ChildUnitNumSearch2.Text & "*'" & _
        "AND DrawDownID=" & Me.DrawDownId & " and GroupID=" & Me.GroupID & " " & _
        "Order By UnitNum,ChildType,CInvoice;"
    Me.DrawnPartsList.RowSource = SQLz
    Call RunningTotal
End Sub

Private Sub ChildVINSearch_Change()
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
        "WHERE VIN Like '*" & Me.ChildVINSearch.Text & "*'" & _
        "AND GroupID=" & Me.GroupID & " " & _
        "Order By VIN,ChildType,CInvoice;"
    Me.OpenPartsList.RowSource = SQLz
End Sub

Private Sub ChildVINSearch2_Change()
    ClearSearchBoxes (Me.ActiveControl.Name)
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE VIN Like '*" & Me.ChildVINSearch2.Text & "*'" & _
        "AND DrawDownID=" & Me.DrawDownId & " and GroupID=" & Me.GroupID & " " & _
        "Order By UnitNum, ChildType,CInvoice;"
    Me.DrawnPartsList.RowSource = SQLz
    Call RunningTotal
End Sub



Private Sub ClearChildType_Click()
    ListBoxClear (Forms!DrawDownManager.ChildPartTypeList)
    If Nz(Me.GroupID, 0) > 0 Then SetChildPartLists (Me.GroupID)
End Sub

Private Sub ClearInvalidDraw_Click()
    If IsNull(Me.DrawDownList.Column(0)) Then Exit Sub
    Dim UnitCnt As Integer
    UnitCnt = Nz(DCount("UnitID", "Units", "DrawDownID=" & Me.DrawDownList.Column(0) & ""), 0)
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Draw: '" & Me.DrawDownList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " UNITS", vbCritical, "Don't Play With Matches!"
        Exit Sub
    End If
    UnitCnt = 1
    UnitCnt = Nz(DCount("UnitRef", "ChildParts", "DrawDownID=" & Me.DrawDownList.Column(0) & ""), 0)
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Draw: '" & Me.DrawDownList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " PARTS", vbCritical, "Stop It, NOW!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Draw: " & Me.DrawDownList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
    If Response = vbYes Then

        SQLz = "Delete * From DrawDowns WHERE DrawDownID =" & Me.DrawDownList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Me.DrawDownList.Requery
End Sub

Private Sub ClearInvoices_Click()
    Call ListBoxClear(Me.Form.Name & "!" & InvoiceList)
End Sub

Private Sub CloseMe_Click()
    Call Form_Load
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub DelScheduleDraws_Click()

    If IsNull(Me.ScheduleList.Column(0)) Then Exit Sub
    Dim iSchID As Long
    Dim Response, FormNameX As String
    Response = MsgBox("Are you SURE that you want to remove All Drawn Units/Parts " & _
                       vbNewLine & "along with ALL of the Draws For Sch: " & Me.ScheduleList.Column(1) & " (All Groups)" & _
                       vbNewLine & vbNewLine & "This cannot be undone!", vbQuestion + vbYesNo + vbCritical, "Are You SURE!?")
    If Response = vbYes Then
        Response = MsgBox("Are you REALLY SURE! that you want to remove All Draws!!" & _
                           vbNewLine & vbNewLine & "This cannot be undone!", vbQuestion + vbYesNo + vbCritical, "THINK!! Are You SURE!?")
        If Response = vbYes Then
            iSchID = Me.ScheduleList.Column(0)
            SQLz = "UPDATE Units Set DrawDownId = Null WHERE UnitID in (SELECT AssetID FROM vw_SixKEys WHERE SchID =" & iSchID & ")"
            DoCmd.SetWarnings False
            MessageUser "Open", "Removing Units From Draws", "just a moment.."
            DoCmd.RunSQL (SQLz)

            SQLz = "UPDATE ChildParts Set DrawDownId = Null WHERE UnitREF in (SELECT AssetID FROM vw_SixKEys WHERE SchID =" & iSchID & ")"
            MessageUser "Update", "Removing Parts From Draws", "just a moment.."
            DoCmd.RunSQL (SQLz)

            SQLz = "DELETE From DrawDowns "
            SQLz = SQLz & " WHERE DrawdownID in ("
            SQLz = SQLz & " SELECT D.DrawDownID FROM DrawDowns as D"
            SQLz = SQLz & " LEFT JOIN Units As U on U.DrawdownID=D.DrawdownID"
            SQLz = SQLz & " WHERE U.UnitID is NULL)"
            MessageUser "Update", "Removing Empty Draws", "just a moment.."
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            MessageUser "Update", "REfreshing Form", "almost done.."
            Me.DrawDownList.Requery
            MessageUser "Close"
        End If
    End If
    Exit Sub
FixMe:
End Sub

Private Sub DrawFunded_Click()
    If FAUserAccessGroup Like "*SYND*" Then
        If Me.DrawDownId & "test" <> "test" Then
AskDate:
            Dim DatePaidOff As Date
            DatePaidOff = InputBox("Provide the DATE Paid Off." & vbNewLine & _
                "          Please enter that date, below." & vbNewLine & _
                "          The Fund Date or Todays is the default." & vbNewLine & vbNewLine & _
                "          (You MUST use a date or 'Cancel')" _
                , "Date Paid Off:", Nz(DLookup("FundDate", "SchGrp", "SGrpid=" & Me.GroupID), Date))
            'If Not IsDate(DatePaidOff) Then GoTo AskDate
            DoCmd.SetWarnings False
            SQLz = "Update Units Set UnitPayOff='" & DatePaidOff & "' Where DrawDownId=" & Me.DrawDownId
            If IsDate(DatePaidOff) Then DoCmd.RunSQL (SQLz)
            SQLz = "Update DrawDowns Set DrawPaidOff='" & DatePaidOff & "' Where DrawDownId=" & Me.DrawDownId
            If IsDate(DatePaidOff) Then DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            Me.Form.Refresh
            Call DrawDownList_Click
        End If
    Else
        Call MsgBox("Sorry only the Syndication Dept" & vbNewLine & "can set the Draw Fund Date", vbCritical, "Don't even rhink about it!")
    End If
End Sub



Private Sub ExcelUnitPartReport_Click()
    If Me.SchID & "test" = "test" Then Exit Sub
    Call MakeDrawDownUnitReport("SchID", Me.SchID, "Unit Draw-Down Report - for Client:" & Me.ClientList.Column(1) & "Sch: " & Me.ScheduleList.Column(1))
End Sub

Private Sub Form_Load()
    OpenUnitsX = 0
    DrawnUnitsX = 0
    Call ClearSearchBoxes
    Call ClearDetails
    Me.Detail.Visible = True
    Me.InvoiceLabel.caption = "Invoices..."
    DoCmd.SetWarnings True
    PreSQLDrawn = "Select AssetID, VIN, UnitNum, Invoice, Format(IFAmount,'$#,#') As IFAmt, UnitVendorPytDate As VendorPaid  From DrawDowns_Units_All "

    Me.SchNum = ""
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.DrawDownList.RowSource = ""
    Me.InvoiceList.RowSource = ""
    Me.AttachedInvoicesList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.OpenUnitsList.RowSource = ""
    Me.DrawnUnitsList.RowSource = ""
    Me.InvoicesInDrawList.RowSource = ""
    Me.ChildInvoiceList.RowSource = ""
    Me.ChildPartTypeList.RowSource = ""
    Me.OpenPartsList.RowSource = ""
    Me.OpenUnitsLabel.caption = ""
    AttachedInvoices_Label.caption = ""
    Me.GroupID = Null
    Me.SchID = Null
    Me.MLID = Null

    SQLx = "UPDATE DISTINCTROW (Units AS U "
    SQLx = SQLx & "INNER JOIN DrawDowns AS D ON U.DrawDownId = D.DrawDownId) "
    SQLx = SQLx & "INNER JOIN DrawDowns ON U.DrawDownId = DrawDowns.DrawDownId "
    SQLx = SQLx & "SET DrawDowns.GroupID = U.GroupREF "
    SQLx = SQLx & "WHERE (((D.groupID)<>U.GroupRef));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLx)
    DoCmd.SetWarnings True

End Sub
Private Sub ClearSearches_Click()
    Call ClearSearchBoxes
    Me.SearchInvoice.SetFocus
    Call SearchInvoice_Change
End Sub

Private Sub ClearSearchBoxes(Optional ExceptSearchBox As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Search" Then
            If Not IsNull(ExceptSearchBox) Then
                If ExceptSearchBox <> Me(Ccrl.Name).Name Then Me(Ccrl.Name) = ""
            Else
                Me(Ccrl.Name) = ""
            End If
        End If
    Next Ccrl
End Sub

Private Function SetListboxes(Typex As String, IDX As Long)

    Dim WhereZ, WhereY, WhereX As String
    Dim SQLy
    Select Case Typex
        Case "Client"
        Case "MLID"
            WhereX = "DrawDowns.MLID=" & IDX & " "
            WhereY = "MLID=" & IDX & ""
            WhereZ = "MLID=" & IDX & ""
        Case "Sch"
            WhereX = "DrawDowns.SchID=" & IDX & " "
            WhereY = "SchID=" & IDX & ""
            WhereZ = "SchID=" & IDX & ""
        Case "Grp"
            WhereX = "DrawDowns.GroupID=" & IDX & " "
            WhereY = "SGrpID=" & IDX & ""
            WhereZ = "GroupID=" & IDX & ""
    End Select
    ListBoxClear (Forms!DrawDownManager.DrawDownList)
    ListBoxClear (Forms!DrawDownManager.InvoiceList)
    If Typex <> "Grp" Then
        Me.DrawDownList.RowSource = ""
        Me.DrawnUnitsList.RowSource = ""
        Me.ChildInvoiceList.RowSource = ""
        Me.OpenUnitsList.RowSource = ""
        Me.InvoiceList.RowSource = ""
        Me.InvoicesInDrawList.RowSource = ""
        Me.AttachedInvoicesList.RowSource = ""
        Me.ChildPartTypeList.RowSource = ""
        Me.OpenPartsList.RowSource = ""
        Me.AddDrawDownBtn.caption = "Add A Draw... Select a Group"
        Me.InvoiceLabel.caption = "Invoices... select a Group"
        Me.OpenUnitsLabel.caption = "Open Units... select a Group"
        Me.DrawnUnitsLabel.caption = "Units in Draw... select a Group"
        Me.SchNum = ""
        Me.GroupID = Null

        ListBoxClear (Forms!DrawDownManager.DrawDownList)

        If Typex = "Client" Then
            '---- The Client List is Clicked ----
            SQLz = "SELECT * From DrawDowns_MstLeases WHERE ClientId =" & IDX & " Order by MLNo;"
            Me.LeaseList.RowSource = SQLz
            Me.ScheduleList.RowSource = ""
            Me.GroupList.RowSource = ""
            Me.MLID = Null
            Me.SchID = Null
            Me.GroupID = Null
            Exit Function
        End If
        If Typex = "MLID" Then
            '---- The Lease List is Clicked ----
            SQLz = "SELECT * FROM DrawDowns_Schedules WHERE MLID =" & IDX & " Order By CDbl(NumInString(Left(Schedule,3))) desc, Schedule desc;"
            Me.ScheduleList.RowSource = SQLz
            Me.ClientName = Me.ClientList.Column(1)
            Me.LeaseName = Me.LeaseList.Column(1)
            Me.GroupList.RowSource = ""
            Me.MLID = Me.LeaseList.Column(0)
            Me.SchID = Null
            Me.GroupID = Null
            Exit Function
        End If
        If Typex = "Sch" Then
            '---- The Schedule List is Clicked ----
            SQLz = "SELECT SGrpID,UnitGroup as [Group],Drawn,Open, Schedule,SchID " & _
                "FROM DrawDowns_Groups " & _
                "WHERE " & WhereY & ";"
            Me.GroupList.RowSource = SQLz
            Me.SchID = IDX
            Me.SchNum = Me.ScheduleList.Column(1)
            Me.SchID = IDX
            Me.GroupID = Null
        End If
    Else
        '---- The Group List is Clicked ----
        Dim CaptionSufix As String
        Dim DrawSchGrpName As String
        Me.GroupID = IDX
        DrawSchGrpName = "Sch " & Me.ScheduleList.Column(1) & " Grp: " & Me.GroupList.Column(1)
        Me.InvoiceLabel.caption = "Invoices for Units in Sch: " & Me.ScheduleList.Column(1) & " Grp: " & Me.GroupList.Column(1)
        If CInt(Me.GroupList.Column(3)) = 0 Then CaptionSufix = "All units Drawn on"
        If CInt(Me.GroupList.Column(3)) = 1 Then CaptionSufix = "1 unit NOT in a Draw"
        If CInt(Me.GroupList.Column(3)) > 1 Then CaptionSufix = "" & Me.GroupList.Column(3) & " units NOT in a Draw"
        Me.OpenUnitsLabel.caption = DrawSchGrpName & " | " & CaptionSufix
        Me.AddDrawDownBtn.caption = "Add A NEW Draw To:" & DrawSchGrpName

        If Me.GroupList.Column(2) = "1" Then CaptionSufix = "" Else CaptionSufix = "s"
        Me.DrawnUnitsLabel.caption = "Unit" & CaptionSufix & " in a Draw " & Me.DrawDownList.Column(3)
        Me.SchNum = DrawSchGrpName
        SQLz = "SELECT * " & _
            "FROM DrawDowns_Invoices_Open_Drawn " & _
            "WHERE " & WhereY & " " & _
            "ORDER BY Invoice, Vendor;"
        Me.InvoiceList.RowSource = SQLz

        SQLz = "SELECT AssetID, VIN, UnitNum, Invoice, Format(VendorCost,'$#,#.#0') As VCost, Format(VendorCostParts,'$#,#.#0') As VCostParts " & _
            "FROM DrawDowns_Units_All As DDA " & _
            "WHERE DrawDownID is null AND " & WhereY & " " & _
            "ORDER By Invoice, AssetId;"
        Me.OpenUnitsList.RowSource = SQLz


        PreSQLDrawn = "Select AssetID, VIN, UnitNum, Invoice, Format(IFAmount,'$#,#.#0') As IFAmt, UnitVendorPytDate As VendorPaid  "
        PreSQLDrawn = PreSQLDrawn & " From DrawDowns_Units_All "
        SQLz = PreSQLDrawn & " WHERE Not DrawDownID is null AND SGrpID=" & IDX & " ORDER By Invoice, AssetId;"
        Me.DrawnUnitsList.RowSource = SQLz

        SQLz = "SELECT * " & _
               "FROM DrawDownsChildInvoices WHERE " & WhereZ
        Me.ChildInvoiceList.RowSource = SQLz

        SQLz = "SELECT * " & _
            "FROM DrawDownChildParts_GroupID " & _
            "WHERE GroupID= " & IDX & " " & _
            "ORDER BY [Type];"
        Me.ChildPartTypeList.RowSource = SQLz

        Call SetChildPartLists(IDX)

        SQLz = "SELECT DrawDowns.DrawDownId, Banks.BankShortName as Bank, DrawDowns.DrawRequestNumber As Request, DrawDowns.DrawNumber As Draw, DrawDowns.DrawName, Count(U.UnitID) AS Units, "
        SQLz = SQLz & "Max(pParts) AS Parts, Format(Sum(UnitIFAmt),'$#,#') AS IfAmount, Format(Max(pCost),'$#,#.#0') AS PartAmt "
        SQLz = SQLz & "FROM (((DrawDowns "
        SQLz = SQLz & "INNER JOIN Banks ON DrawDowns.DrawBankId = Banks.BankId) "
        SQLz = SQLz & "LEFT JOIN Units As U ON DrawDowns.DrawDownId = U.DrawDownId) "
        SQLz = SQLz & "LEFT OUTER JOIN (Select DrawDownId, Count(ChildPartID) as pParts, Sum(ChildFinalCost) as pCost "
        SQLz = SQLz & "From ChildParts Group By DrawDownId) As CP "
        SQLz = SQLz & "ON DrawDowns.DrawDownId = CP.DrawDownId) "
        SQLz = SQLz & "WHERE " & WhereX
        SQLz = SQLz & "GROUP BY DrawDowns.DrawDownId, Banks.BankShortName, DrawDowns.DrawRequestNumber, DrawDowns.DrawNumber, DrawDowns.DrawName "
        SQLz = SQLz & "ORDER BY DrawDowns.DrawName, DrawDowns.DrawNumber;"
        Me.DrawDownList.RowSource = SQLz
        Call ClearDetails
        Call UpdateListBoxes
        Call GetUnitCounts
    End If
    InvoicesInDrawList.RowSource = ""
    Me.AttachedInvoicesList.RowSource = ""
    Me.AttachedInvoices_Label.caption = "Sch/Grps that also have Invoice... select a Invoice "
End Function
Private Function SetChildPartLists(GroupIDx As Long)
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
        "WHERE GroupID=" & Me.GroupID & " " & _
        "Order By AssetID,ChildType,CInvoice;"
    Me.OpenPartsList.RowSource = SQLz

    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE DrawDownID=" & Me.DrawDownId & " and GroupID=" & Me.GroupID & " " & _
        "Order By AssetID,ChildType,CInvoice;"
    Me.DrawnPartsList.RowSource = SQLz


End Function

Private Sub ClientList_Click()
    Call SetListboxes("Client", Me.ClientList.Column(0))
End Sub

Private Sub Leaselist_Click()
    Call SetListboxes("MLID", Me.LeaseList.Column(0))
End Sub

Private Sub MarkVendorPaid_Click()
    If [DrawnUnitsList].[ItemsSelected].[Count] < 1 Then
        Call MsgBox("You MUST Select Unit(s) First", vbCritical, "No Units Selected")
        Exit Sub
    End If
    Dim VendorPaidOff As String
    On Error Resume Next
    VendorPaidOff = InputBox("Provide the DATE The VENDOR was PAID." & vbNewLine & _
        "          Please enter that date, below." & vbNewLine & _
        "          The Paid Date or Todays is the default." & vbNewLine & vbNewLine & _
        "          (You MUST use a date or 'Cancel')" & vbNewLine & vbNewLine & _
        "          Note: This will change any date already entered!", "Enter a Paid Date")

    On Error GoTo 0
    If Not IsDate(VendorPaidOff) Then Exit Sub
    DoCmd.SetWarnings False
    Dim InvDatex, InvPaidDatex As String
    Dim DrawnDowni As Long
    For Each ItemVar In Me.DrawnUnitsList.ItemsSelected
        InvX = DLookup("UnitInvoiceNum", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar))
        DrawnDowni = Nz(DLookup("DrawDownId", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar)), 0)

        SQLy = "Update Units Set UnitVendorPytDate='" & VendorPaidOff & "' "
        SQLy = SQLy & "Where UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar)
        DoCmd.RunSQL (SQLy)

        SQLz = "Update ChildParts Set ChildVendorPytDate='" & VendorPaidOff & "' "
        SQLz = SQLz & " Where UNitRef=" & Me.DrawnUnitsList.ItemData(ItemVar)
        SQLz = SQLz & " AND DrawDownID=" & DrawnDowni & " "
        DoCmd.RunSQL (SQLz)
    Next ItemVar
    If VendorPaidOff & "test" <> "test" Then Call MsgBox("Vendor Paid Date set", vbInformation, "Vendor Marked Paid (w/parts)")
    Me.DrawnPartsList.Requery
    DoCmd.SetWarnings True
End Sub

Private Sub OpenMassMod_Click()
    If Me.GroupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    AddReportLog "OpenMassMod_Click"
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = Me.ClientList.Column(0)
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = Me.LeaseList.Column(0)
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = Me.SchID
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = Me.GroupID
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
End Sub

Private Sub ScheduleList_Click()
    Call SetListboxes("Sch", Me.ScheduleList.Column(0))
End Sub

Private Sub GroupList_Click()
    Call SetListboxes("Grp", Me.GroupList.Column(0))
End Sub
Private Sub DrawDownList_Click()
    Call ClearDetails

    If Nz(Me.DrawDownList.Column(0), 0) = 0 Then Exit Sub

    Dim DrawX  As Integer
    Me.DrawnUnitCost = 0
    Me.DrawSelectedCost = 0
    Me.DrawnSelectedCostWParts = 0
    DrawX = Me.DrawDownList.Column(0)

    SQLz = "SELECT Invoice, Units, TCost As Amount, DrawDownID, SchID "
    SQLz = SQLz & " FROM DrawDowns_Invoices_Drawn "
    SQLz = SQLz & " WHERE DrawDownID=" & Me.DrawDownList.Column(0) & ";"
    Me.InvoicesInDrawList.RowSource = SQLz

    SQLz = PreSQLDrawn & " WHERE DrawDownID=" & Me.DrawDownList & " Order By Invoice, AssetId;"
    Me.DrawnUnitsList.RowSource = SQLz
    SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
           "WHERE DrawDownID=" & Me.DrawDownList & " " & _
           "ORDER BY CInvoice,AssetID;"
    Me.DrawnPartsList.RowSource = SQLz

    SQLz = "Select * From vw_DrawDowns_Invoices Where DrawDownID=" & DrawX
    Set RST = CurrentDb.OpenRecordset(SQLz)
    If RST.RecordCount >= 1 Then RST.MoveFirst
    Me.DrawID = Me.DrawDownList.Column(0)
    Me.iBank = Me.DrawDownList.Column(1)
    Me.RequestNumber = Me.DrawDownList.Column(2)
    Me.DrawnUnitsLabel.caption = "Units in " & Me.DrawDownList.Column(4)
    Me.UnitDrawCount = "Units:" & Me.DrawDownList.Column(5)
    Me.ChildDrawCount = "Parts:" & Me.DrawDownList.Column(6)
    If RST.RecordCount >= 1 Then
        If Nz(RST!Funded, 0) = 0 Then Me.fStatus = "NOT Funded" Else Me.fStatus = "Funded"
    End If
    If RST.RecordCount > 0 Then Me.TotalAmountDrawn = CCur(Nz(RST!DrawnAmount, 0)) Else Me.TotalAmountDrawn = 0
    RST.close
    SQLz = "Select * From DrawDowns Where DrawDownID=" & DrawX
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        Me.DrawDownId = DrawX
        Me.DrawNumber = RST!DrawNumber
        Me.DrawBankAccountNumber = RST!DrawBankAccountNumber
        Me.DateDrawn = RST!DateDrawn
        Me.PPCertSent = RST!PPCertSent
        Me.DrawPaidOff = RST!DrawPaidOff
        Me.MLID = Me.LeaseList
        Me.DrawCreditLimit = RST!DrawCreditLimit
        Me.GroupFunded = DLookup("FundDate", "SchGrp", "SGrpid=" & Me.GroupID)
        RST.close
    End If
    Set RST = Nothing
    Me.AvailAmt = Nz(Me.DrawCreditLimit, 0) - Nz(Me.TotalAmountDrawn, 0)
    'Me.TotalAmountDrawn = CCur(Nz(DSum("UnitCost", "DrawDowns_Units_ALL", "RequestNumber='" & Me.RequestNumber & "'"), 0) + Nz(DSum("CCost", "DrawDowns_ChildParts_Drawn", "DrawRequestNumber='" & Me.RequestNumber & "'"), 0))
End Sub
Private Sub ClearDetails()
    Dim ctrl   As Control
    For Each ctl In Me.Form
        If Me(ctl.Name).Tag = "RequestDetails" Then Me(ctl.Name) = ""
    Next ctl
    Me.UnitDrawCount.Requery
    Me.ChildDrawCount.Requery
End Sub
Private Sub InvoicesInDrawList_Click()
    If Me.InvoicesInDrawList.ListCount < 1 Then Exit Sub
    Me.ChildInvoiceList = Me.InvoicesInDrawList.Column(0)
    Call SelectUnits_Drawn(Me.InvoicesInDrawList.Column(0), Me.InvoicesInDrawList.Column(3))
    Call SelectUnits_Open(Me.InvoicesInDrawList.Column(0))
    Call SelectParts_Drawn(Me.InvoicesInDrawList.Column(0))
    Call SelectParts_Open(Me.InvoicesInDrawList.Column(0))
    ListBoxClear (Forms!DrawDownManager.InvoiceList)
    ListBoxClear (Forms!DrawDownManager.AttachedInvoicesList)
    Me.ScheduleList = Me.InvoicesInDrawList.Column(4)
    SQLz = "Select * From DrawDowns_Invoices_Open_Drawn Where SchId =" & Me.ScheduleList & " ORDER BY Vendor, Invoice;"
    Me.InvoiceList.RowSource = SQLz
    Me.InvoiceList = Me.InvoicesInDrawList.Column(0)
    SQLz = "Select * From DrawDowns_AttachedInvoices " & _
        "WHERE Invoice='" & Me.InvoicesInDrawList.Column(0) & "' Order By [DrawName],Draw;"
    Me.AttachedInvoicesList.RowSource = SQLz
    Me.AttachedInvoicesList = Me.InvoicesInDrawList.Column(3)

    Me.InvoiceLabel.caption = "Invoices for Units in Schedule " & Me.ScheduleList.Column(1)
    Me.OpenUnitsLabel.caption = "Sch " & Me.ScheduleList.Column(1) & " " & [OpenUnitsList].[ListCount] - 1 & " Units not in a Draw"
End Sub
Private Sub ChildInvoiceList_Click()
    If Me.ChildInvoiceList.ListCount < 1 Then Exit Sub
    Call SelectParts_Open(Me.ChildInvoiceList.Column(0))
    Call SelectParts_Drawn(Me.ChildInvoiceList.Column(0))
    Me.InvoicesInDrawList = Me.ChildInvoiceList
End Sub
Private Function SelectUnits_Drawn(InvoiceNum As String, DrawID As Integer)
    Dim RST    As DAO.Recordset
    Dim InvCount, x, Y As Integer
    ListBoxClear (Forms!DrawDownManager.DrawnUnitsList)
    SQLz = "Select AssetID FROM DrawDowns_Units_All WHERE Not DrawDownID Is Null AND Invoice='" & InvoiceNum & "' AND DrawDownID=" & DrawID & " Order By Invoice, AssetId;"
    Set RST = CurrentDb.OpenRecordset(SQLz)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo Cleanup
    InvCount = RST.RecordCount
    If InvCount = 0 Then GoTo Cleanup
    RST.MoveFirst
    For x = 1 To InvCount
        For Y = 0 To Me.DrawnUnitsList.ListCount
            If (Me.DrawnUnitsList.ItemData(Y) = CStr(RST!AssetID)) Then Me.DrawnUnitsList.Selected(Y) = True
        Next Y
        RST.MoveNext
    Next x
Cleanup:
    RST.close
    Set RST = Nothing
    Exit Function
End Function
Private Function SelectParts_Open(InvoiceNum As String)
    Dim RST    As DAO.Recordset
    Dim InvCount, x, Y As Integer
    ListBoxClear (Forms!DrawDownManager.OpenPartsList)
    SQLz = "Select ChildPartID FROM DrawDowns_ChildParts_Open " & _
        "WHERE GroupID=" & Me.GroupID & " " & _
        "AND CInvoice='" & InvoiceNum & "' " & _
        "Order By CInvoice,AssetID,ChildType;"
    '  "AND ChildType='" & Me.ChildPartTypeList.Column(5) & "' " & _

       Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenDynaset, dbSeeChanges)
       On Error Resume Next
       RST.MoveFirst
       RST.MoveLast
       On Error GoTo Cleanup
       InvCount = RST.RecordCount
       If InvCount = 0 Then GoTo Cleanup
       RST.MoveFirst
       For x = 1 To InvCount
           For Y = 0 To Me.OpenPartsList.ListCount
               If (Me.OpenPartsList.ItemData(Y) = CStr(RST!ChildPartID)) Then Me.OpenPartsList.Selected(Y) = True
           Next Y
           RST.MoveNext
       Next x
Cleanup:
       RST.close
       Set RST = Nothing
       Exit Function
End Function
Private Function SelectParts_Drawn(InvoiceNum As String)
    Dim RST    As DAO.Recordset
    Dim InvCount, x, Y As Integer
    ListBoxClear (Forms!DrawDownManager.DrawnPartsList)
    SQLz = "Select ChildPartID FROM DrawDowns_ChildParts_Drawn " & _
        "WHERE CInvoice='" & InvoiceNum & "' " & _
        "Order By CInvoice,AssetID,ChildType;"
    '  "AND ChildType='" & Me.ChildPartTypeList.Column(5) & "' " & _

       Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenDynaset, dbSeeChanges)
       On Error Resume Next
       RST.MoveFirst
       RST.MoveLast
       On Error GoTo Cleanup
       InvCount = RST.RecordCount
       If InvCount = 0 Then GoTo Cleanup
       RST.MoveFirst
       For x = 1 To InvCount
           For Y = 0 To Me.DrawnPartsList.ListCount
               If (Me.DrawnPartsList.ItemData(Y) = CStr(RST!ChildPartID)) Then Me.DrawnPartsList.Selected(Y) = True
           Next Y
           RST.MoveNext
       Next x
Cleanup:
       RST.close
       Set RST = Nothing
       Exit Function
End Function
Private Sub PickModeToggle_Click()
    Dim i As Integer
    For i = Me.OpenUnitsList.ListCount - 1 To 0 Step -1
        Me.OpenUnitsList.Selected(i) = 0
    Next i
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Pick_DrawUnits_WorkTable")
    DoCmd.SetWarnings True
    If Me.PickModeToggle = True Then
        Me.PickModeToggle.caption = ChrW(9660) & "Pick On"
    Else
        Me.PickModeToggle.caption = "Pick Off"
    End If
    Call GetUnitCounts
End Sub
Private Function SelectUnits_Open(InvoiceNum As String)
    Dim RST    As DAO.Recordset
    Dim InvCount, x, Y As Integer
    ListBoxClear (Forms!DrawDownManager.OpenUnitsList)
    SQLz = "Select AssetID FROM DrawDowns_Units_All WHERE DrawDownID Is Null AND Invoice='" & InvoiceNum & "' Order By Invoice, AssetId;"
    Set RST = CurrentDb.OpenRecordset(SQLz)
    On Error Resume Next
    RST.MoveFirst
    RST.MoveLast
    On Error GoTo Cleanup
    InvCount = RST.RecordCount
    If InvCount = 0 Then GoTo Cleanup
    RST.MoveFirst
    For x = 1 To InvCount
        For Y = 0 To Me.OpenUnitsList.ListCount
            If (Me.OpenUnitsList.ItemData(Y) = CStr(RST!AssetID)) Then Me.OpenUnitsList.Selected(Y) = True
        Next Y
        RST.MoveNext
    Next x
Cleanup:
    RST.close
    Set RST = Nothing
    Exit Function
End Function
Private Sub OpenUnitsList_Click()
    If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.OpenUnitsList.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Pick_DrawUnits_WorkTable(UnitID) SELECT " & Me.OpenUnitsList.Column(0) & ";")
        DoCmd.SetWarnings True
        Call ReselectPickedUnits
    End If
    Call GetUnitCounts
End Sub
Public Sub ReselectPickedUnits()
    Dim i, x As Integer
    Dim pUnits As String
    Dim RST As DAO.Recordset
    SQLz = "Select * From Pick_DrawUnits_WorkTable Order by UnitID Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For x = 1 To RST.RecordCount
            For i = Me.OpenUnitsList.ListCount - 1 To 0 Step -1
                If Me.OpenUnitsList.ItemData(i) = CStr(RST!UnitID) Then Me.OpenUnitsList.Selected(i) = 1
            Next i
            RST.MoveNext
        Next x
    End If
    RST.close
    Set RST = Nothing
End Sub
Private Sub DrawnUnitsList_Click()
    Call GetUnitCounts
End Sub
Private Sub OpenPartsList_Click()
    Call GetUnitCounts
End Sub
Private Sub DrawnPartsList_Click()
    Call GetUnitCounts
End Sub
Private Sub InvoiceList_Click()

    InvX = Me.InvoiceList.Column(0)

    If Me.InvoiceList.Column(3) <> 0 Then
        If Me.OpenUnitsList.ListCount > 1 Then
            RSx = Split(Me.OpenUnitsList.RowSource, "Order By")(0)
            Me.OpenUnitsList.RowSource = RSx & " Order By Invoice, AssetID;"
        End If
    End If
    ListBoxClear (Forms!DrawDownManager.InvoicesInDrawList)
    ListBoxClear (Forms!DrawDownManager.DrawnUnitsList)
    ListBoxClear (Forms!DrawDownManager.AttachedInvoicesList)
    Me.ScheduleList = Me.InvoiceList.Column(11)
    If Me.InvoiceList.Column(8) <> "" Then
        ListBoxClear (Forms!DrawDownManager.DrawDownList)
        Me.DrawDownList = Me.InvoiceList.Column(8)
        Call DrawDownList_Click
    End If
    Me.InvoicesInDrawList = InvX

    SQLz = "Select * From DrawDowns_AttachedInvoices " & _
        "WHERE Invoice='" & InvX & "' Order By [DrawName];"
    Me.AttachedInvoicesList.RowSource = SQLz
    Call SelectUnits_Open(InvX)

    If Me.InvoiceList.Column(7) <> "" Then
        Call SelectUnits_Drawn(InvX, Me.InvoiceList.Column(7))
        Me.AttachedInvoicesList = Me.InvoiceList.Column(7)
        Me.AttachedInvoices_Label.caption = "Sch/Grps that also have Invoice: " & Me.InvoiceList.Column(0)
    End If
    If Me.InvoiceList.Column(7) > 0 Then Me.DrawDownList = Me.InvoiceList.Column(7)


    SQLz = "Select AssetID, VIN, UnitNum As [Unit#], Invoice, Format(UnitCost,'$#,#') As Cost From DrawDowns_Units_All " & _
        "WHERE DrawDownID Is Null AND Invoice='" & InvX & "' Order By AssetId;"
    '  Me.OpenUnitsList.RowSource = SQLz
    Call ClearSearchBoxes
    Call GetUnitCounts
End Sub

Private Sub AttachedInvoicesList_DblClick(Cancel As Integer)
    If Me.DrawID <> Me.AttachedInvoicesList.Column(0) Then
        Committ = MsgBox("Do you want to Jump to Draw: " & _
            vbNewLine & vbNewLine & "'" & Me.AttachedInvoicesList.Column(3) & "'" & _
            vbNewLine & vbNewLine & "This will clear your selection", vbInformation + vbYesNo, "Look before you Jump")
        If Committ = vbNo Then Exit Sub
        Call MoveToDraw(Me.AttachedInvoicesList.Column(0), Me.AttachedInvoicesList.Column(6))
    End If
End Sub
Public Function MoveToDraw(DrawIDx As Long, Optional Invoicex As String)
    Dim RST    As DAO.Recordset
    Dim InvCount, x, Y As Integer
    ListBoxClear (Forms!DrawDownManager.OpenUnitsList)
    SQLz = "Select * FROM DrawDowns WHERE DrawDownID=" & DrawIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.ClientList = RST!ClientID
    Call ClientList_Click
    Me.LeaseList = RST!MLID
    Call SetListboxes("MLID", RST!MLID)
    Me.ScheduleList = RST!SchID
    Call SetListboxes("Sch", RST!SchID)
    Me.GroupList = RST!GroupID
    Call SetListboxes("Grp", RST!GroupID)
    Me.DrawDownList = RST!DrawDownId
    Call DrawDownList_Click
    If Invoicex <> "" Then
        Me.InvoiceList = Invoicex
        Call InvoiceList_Click
    End If
    Call InvoicesInDrawList_Click
    RST.close
    Set RST = Nothing
    MsgBox "You now viewing " & Me.SchNum & "", , "Whew! Made It"
End Function
Private Sub AttachedInvoicesList_Click()
    Me.DrawDownList = Me.AttachedInvoicesList.Column(0)
    Me.InvoicesInDrawList = Me.AttachedInvoicesList.Column(6)
    Call DrawDownList_Click
    Call SelectUnits_Drawn(Me.AttachedInvoicesList.Column(6), Me.AttachedInvoicesList.Column(0))
    Call ClearSearchBoxes
End Sub
Private Sub AddDrawDownBtn_Click()
    If Me.GroupID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "DrawDownAddaDraw"
    Call DoCmd.OpenForm("DrawDownAddaDraw", , , , , , "Add>" & Me.GroupID)
End Sub

Private Sub EditMe_Click()
    If Me.DrawDownList & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "DrawDownAddaDraw"
    Call DoCmd.OpenForm("DrawDownAddaDraw", , , , , , "Edit>" & Me.DrawDownList)
End Sub
Private Sub DeleteDraw_Click()
    If IsNull(Me.DrawDownList.Column(0)) Then Exit Sub
    Dim UnitCnt As Integer
    UnitCnt = Nz(DCount("UnitID", "Units", "DrawDownID=" & Me.DrawDownList.Column(0) & ""), 0)
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Draw: '" & Me.DrawDownList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " UNITS", vbCritical, "Don't Play With Matches!"
        Exit Sub
    End If
    UnitCnt = 1
    UnitCnt = Nz(DCount("UnitRef", "ChildParts", "DrawDownID=" & Me.DrawDownList.Column(0) & ""), 0)
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Draw: '" & Me.DrawDownList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " PARTS", vbCritical, "Stop It, NOW!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Draw: " & Me.DrawDownList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
    If Response = vbYes Then

        SQLz = "Delete * From DrawDowns WHERE DrawDownID =" & Me.DrawDownList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Me.DrawDownList.Requery
End Sub
Private Sub MarkUpdate()
    Me.ModifiedBy = Forms!FAMainMenu.USERID
    Me.ModifiedOn = Now()
End Sub

Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From DrawUnits_WorkTable;")
    DoCmd.SetWarnings True
End Sub

Private Sub GetAll_Click()
    For i = 0 To OpenUnitsList.ListCount - 1
        OpenUnitsList.Selected(i) = True
    Next i
    GetUnitCounts
End Sub
Private Sub GetAllOpenParts_Click()
    For i = 0 To OpenPartsList.ListCount - 1
        OpenPartsList.Selected(i) = True
    Next i
    GetUnitCounts
End Sub
Private Sub GetAllDrawnParts_Click()
    For i = 0 To DrawnPartsList.ListCount - 1
        DrawnPartsList.Selected(i) = True
    Next i
    GetUnitCounts
End Sub

Private Sub AddParts_Click()
    Call AddPartsToDraw(Me.DrawID)
End Sub
Private Function AddPartsToDraw(DrawIDx As Long, Optional CopyInvToParts As Boolean, Optional InvX As String)
    If Me.DrawID & "test" = "test" Then
        MsgBox "Please Select a Draw", vbCritical, "Pick A Draw... Any Draw"
        Exit Function
    End If
    If IsDate(Me.DrawPaidOff) Then
        Committ = MsgBox("This Draw is Paid Off as of: " & DrawPaidOff & "!" & _
            vbNewLine & "Do you STLL want to add Parts?" & _
            vbNewLine & vbNewLine & "Yes/No", vbInformation + vbYesNo, "Look before you Jump")
        If Committ = vbNo Then Exit Function
    End If
    Dim ItemVar As Variant
    For Each ItemVar In Me.OpenPartsList.ItemsSelected
        SQLz = "Update ChildParts Set "
        SQLz = SQLz & " DrawDownId=" & DrawIDx & " "
        SQLz = SQLz & ",cIFAmount = NZ(childFinalCost,0) "
        SQLz = SQLz & ",cIFDate = '" & Date & "' "
        SQLz = SQLz & ",cIFBankID = IIf(nz(cIFBankID,0),0,5649) "
        SQLz = SQLz & " Where ChildPartID=" & Me.OpenPartsList.ItemData(ItemVar) & ";"
        '  If CopyInvToParts = True Then SQLz = SQLz & "ChildInvoiceNumber='" & InvX & "' "
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next ItemVar

    On Error GoTo fixme1
    Call UpdateListBoxes
    Call SetChildPartLists(Me.GroupID)
    Exit Function
fixme1:
End Function
Private Sub RemoveParts_Click()
    If Me.DrawID & "test" = "test" Then
        MsgBox "Please Select a Draw", vbCritical, "Need A Draw"
        Exit Sub
    End If
    If IsDate(Me.DrawPaidOff) Then
        Committ = MsgBox("This Draw is Paid Off as of: " & DrawPaidOff & "!" & _
            vbNewLine & "Do you STLL want to Remove Parts?" & _
            vbNewLine & vbNewLine & "Yes/No", vbInformation + vbYesNo, "Look before you Jump")
        If Committ = vbNo Then Exit Sub
    End If
    Committ = MsgBox("Do you want to Remove: " & DrawnSelectedPartsCount & "Parts From Draw: " & _
        vbNewLine & vbNewLine & "'" & Me.AttachedInvoicesList.Column(3) & "'" & _
        vbNewLine & vbNewLine & "This cannot not be reversed", vbInformation + vbYesNo, "Look before you Jump")
    If Committ = vbNo Then Exit Sub
    Dim ItemVar As Variant
    For Each ItemVar In Me.DrawnPartsList.ItemsSelected
        SQLz = "UPDATE ChildParts Set DrawDownId=NULL "
        SQLz = SQLz & ",cIFAmount = NULL "
        SQLz = SQLz & ",cIFDate = NULL "
        SQLz = SQLz & ",cIFBankID = NULL "
        SQLz = SQLz & " WHERE ChildPartID=" & Me.DrawnPartsList.ItemData(ItemVar) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next ItemVar
    On Error GoTo fixme1
    Call UpdateListBoxes
    Call SetChildPartLists(Me.GroupID)
    Me.ChildInvoiceList.Requery
    Exit Sub
fixme1:
End Sub

Private Sub AddUnits_Click()
    If Me.DrawID & "test" = "test" Then
        MsgBox "Please Select a Draw", vbCritical, "Need A Draw"
        Exit Sub
    End If
    If IsDate(Me.DrawPaidOff) Then
        Committ = MsgBox("This Draw is Paid Off as of: " & DrawPaidOff & "!" & _
            vbNewLine & "Do you STLL want to add Units?" & _
            vbNewLine & vbNewLine & "Yes/No", vbInformation + vbYesNo, "Look before you Jump")
        If Committ = vbNo Then Exit Sub
    End If
    Dim ItemVar As Variant
    Dim CostToUse, SQLy As String
    Dim GetParts, CopyInvToParts As Boolean
    Dim DrawX  As Integer
    Dim iBankx As String
    Dim UnitIDx As Long
    Dim InvDatex, InvPaidDatex As String
    Dim IFBankIDi, VendorIDi As Long

    If CostToUseFrame = 1 Then CostToUse = "VendorCost"
    If CostToUseFrame = 2 Then CostToUse = "VendorCostParts"
    If CostToUseFrame = 3 Then CostToUse = "VendorCostParts"
    If CostToUseFrame > 1 Then GetParts = True
    If CostToUseFrame = 3 Then CopyInvToParts = True
    iBankx = Me.DrawDownList.Column(1)
    DrawX = Me.DrawID

    For Each ItemVar In Me.OpenUnitsList.ItemsSelected
        UnitIDx = Me.OpenUnitsList.ItemData(ItemVar)
        InvX = DLookup("UnitInvoiceNum", "Units", "UnitId=" & UnitIDx)
        InvDatex = Nz(DLookup("UnitInvoiceDate", "Units", "UnitId=" & UnitIDx), "")
        InvPaidDatex = Nz(DLookup("UnitVendorPytDate", "Units", "UnitId=" & UnitIDx), "")
        IFBankIDi = Nz(DLookup("DrawBankId", "DrawDowns", "DrawDownID=" & Me.DrawID & ""), 5649)
        VendorIDi = Nz(DLookup("VendorRef", "Units", "UnitId=" & UnitIDx), 0)

        InvX = DLookup("Invoice", "DrawDowns_Units_All", "AssetID=" & UnitIDx)
        SQLz = "UPDATE Units SET Units.DrawDownID =" & Me.DrawID & " "
        SQLz = SQLz & ",UnitIFBankID = " & IFBankIDi
        SQLz = SQLz & ",UnitIFAmt  = IIf(Nz(VendorCost, 0) > 0, VendorCost, UnitFinalCost)"
        'SQLz = SQLz & ",UnitIFDate = '" & Date & "' "
        SQLz = SQLz & " WHERE UnitID=" & UnitIDx & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        If GetParts = True Then
            SQLy = "UPDATE ChildParts Set ChildInvoiceNum=Nz(ChildInvoiceNum,'" & InvX & "')"
            SQLy = SQLy & " ,ChildInvdt =Nz(ChildInvdt,'" & InvDatex & "') "
            SQLy = SQLy & " ,ChildVendorPytDate = Nz(ChildVendorPytDate,'" & InvPaidDatex & "') "
            SQLy = SQLy & " ,DrawDownID =" & Me.DrawID & " "
            SQLy = SQLy & " ,cIFAmount = childFinalCost "
            'SQLy = SQLy & " ,cIFDate = '" & Date & "' "
            SQLy = SQLy & " ,cIFBankID=" & Nz(IFBankIDi, 5649) & " "
            SQLy = SQLy & " WHERE UnitRef=" & UnitIDx
            SQLy = SQLy & " AND ChildVendorREF = " & VendorIDi & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLy)
            DoCmd.SetWarnings True
        End If
    Next ItemVar

    On Error GoTo fixme1
    Me.InvoiceList = InvX
    Me.DrawID = DrawX
    Call UpdateListBoxes
    Call GroupList_Click
    Call DrawDownList_Click
    If Me.InvoiceList & "test" <> "test" Then Call InvoiceList_Click
fixme1:
    Exit Sub
End Sub
Private Sub RemoveGetAll_Click()
    For i = 0 To DrawnUnitsList.ListCount - 1
        DrawnUnitsList.Selected(i) = True
    Next i
    GetUnitCounts
End Sub
Private Sub RemoveUnits_Click()
    Dim ItemVar As Variant
    If IsDate(Me.DrawPaidOff) Then
        Committ = MsgBox("This Draw is Paid Off as of: " & DrawPaidOff & "!" & _
            vbNewLine & "Do you STLL want to REMOVE Units?" & _
            vbNewLine & vbNewLine & "Yes/No", vbInformation + vbYesNo, "Look before you Jump")
        If Committ = vbNo Then Exit Sub
    End If
    For Each ItemVar In Me.DrawnUnitsList.ItemsSelected
        SQLz = "UPDATE Units SET DrawDownId = Null "
        SQLz = SQLz & ", UnitIFAmt = Null "
        SQLz = SQLz & ", UnitIFDate = Null "
        SQLz = SQLz & ", UnitIFBankID= Null "
        SQLz = SQLz & "WHERE UnitID =" & Me.DrawnUnitsList.ItemData(ItemVar) & ";"

        'SQLy = "UPDATE ChildParts Set DrawDownID=Null "
        'SQLy = SQLy & "WHERE UnitRef=" & Me.DrawnUnitsList.ItemData(ItemVar) & ";"

        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        'DoCmd.RunSQL (SQLy)
        DoCmd.SetWarnings True
    Next ItemVar
    Call UpdateListBoxes
    Call GroupList_Click
    Call DrawDownList_Click
    If Me.InvoiceList & "test" <> "test" Then Call InvoiceList_Click
End Sub

Private Sub UpdateListBoxes()
    Me.LeaseList.Requery
    Me.OpenUnitsList.Requery
    Me.ScheduleList.Requery
    Me.OpenUnitsList.Requery
    Me.DrawnUnitsList.Requery
    Me.DrawDownList.Requery
    Me.DrawnPartsList.Requery
    Me.InvoiceList.Requery
    Call RunningTotal
End Sub

Public Function SearchOpenUnitsList(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetID = ""
    If sField <> "VIN" Then Me.SearchVin = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum = ""
    If sField <> "Invoice" Then Me.SearchInvoice = ""
    If sField = "ClearAll" Then Exit Function
    'If Len(Me("Search" & sField).Text) >= 1 And Len(Me("Search" & sField).Text) < 3 Then Exit Function
    SQLz = "SELECT AssetID, VIN, UnitNum, Invoice, Format(OEC,'$#,#') As Cost " & _
        "FROM DrawDowns_Units_All " & _
        "WHERE DrawDownID Is Null AND SchId=" & Me.ScheduleList & " AND " & sField & " LIKE '*" & Me("Search" & sField).Text & "*' "
    If Me.GroupID <> "" Then SQLz = SQLz & " AND SGrpID =" & Me.GroupID & " "
    SQLz = SQLz & " ORDER BY " & sField & ";"
    Me.OpenUnitsList.RowSource = SQLz
    Call RunningTotal
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Function

Private Sub SearchAssetId_Change()
    Call SearchOpenUnitsList("AssetID")
End Sub

Private Sub SearchVIN_Change()
    Call SearchOpenUnitsList("VIN")
End Sub

Private Sub SearchUnitNum_Change()
    Call SearchOpenUnitsList("UnitNum")
End Sub

Private Sub SearchInvoice_Change()
    Call SearchOpenUnitsList("Invoice")
End Sub

Private Sub SearchAvInvoices_Change()
    SQLz = "SELECT * FROM DrawableUnitInvoices WHERE SchID=" & Me.SchID & " And Invoice Like '*" & Me.SearchAvInvoices.Text & "*';"
    Me.InvoiceList.RowSource = SQLz
End Sub
Public Function SearchDrawnUnitsList(sField As String)
    If sField <> "AssetID" Then Me.Search2AssetID = ""
    If sField <> "VIN" Then Me.Search2VIN = ""
    If sField <> "UnitNum" Then Me.Search2UnitNum = ""
    If sField <> "Invoice" Then Me.Search2Invoice = ""
    If sField <> "UnitVendorPytDate" Then Me.Search2UnitVendorPytDate = ""
    If sField = "ClearAll" Then Exit Function
    SQLz = PreSQLDrawn & " " & _
        "WHERE NOT DrawDownID Is Null AND SchId=" & Me.ScheduleList & " AND " & sField & " LIKE '*" & Me("Search2" & sField).Text & "*' "
    'PreSQLDrawn = "Select AssetID, VIN, UnitNum, Invoice, Format(IFAmount,'$#,#') As IFAmt, UnitVendorPytDate As VendorPaid  From DrawDowns_Units_All "
    'SQLz = PreSQLDrawn & " WHERE Not DrawDownID is null AND SchId=" & Me.ScheduleList & " AND AssetID Like '*" & Me.SearchAsset2.Text & "*' Order By AssetId;"
    SQLz = SQLz & " ORDER BY " & sField & ";"
    Me.DrawnUnitsList.RowSource = SQLz
    Call RunningTotal
    'If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Function

Private Sub Search2AssetID_Change()
    Call SearchDrawnUnitsList("AssetID")
End Sub

Private Sub Search2VIN_Change()
    Call SearchDrawnUnitsList("VIN")
End Sub

Private Sub Search2UnitNum_Change()
    Call SearchDrawnUnitsList("UnitNum")
End Sub

Private Sub Search2Invoice_Change()
    Call SearchDrawnUnitsList("Invoice")
End Sub

Private Sub SearchUnitVendorPytDate_Change()
    Call SearchDrawnUnitsList("UnitVendorPytDate")
End Sub
Public Sub RunningTotal()
    Dim TotalCost As Currency
    Dim SelectedCost As Currency
    Dim SelectedCostWParts As Currency
    Dim DrawnCostWParts As Currency
    TotalCost = 0
    SelectedCost = 0
    SelectedCostWParts = 0
    DrawnCostWParts = 0
    Me.DrawnUnitCost = 0
    Me.DrawSelectedCost = 0
    Me.DrawnSelectedCostWParts = 0

    If DrawnUnitsList.ListCount > 0 Then
        Dim iRow As Integer
        For iRow = 0 To (DrawnUnitsList.ListCount)
            If Nz(DrawnUnitsList.Column(4, iRow), 0) <> "$" Then TotalCost = TotalCost + Nz(DrawnUnitsList.Column(4, iRow), 0)
            If DrawnUnitsList.Selected(iRow) = True And Nz(DrawnUnitsList.Column(4, iRow), 0) <> "$" Then
                SelectedCost = SelectedCost + Nz(DrawnUnitsList.Column(4, iRow), 0)
                If Nz(Me.DrawDownId, "") = "" Then
                    DrawnCostWParts = DrawnCostWParts + Nz(DSum("cIFAmount", "ChildParts", "UnitREF=" & Nz(DrawnUnitsList.Column(0, iRow), 0) & ""), 0)
                Else
                    DrawnCostWParts = DrawnCostWParts + Nz(DSum("cIFAmount", "ChildParts", "DrawDownId=" & Me.DrawDownId & " AND UnitREF=" & Nz(DrawnUnitsList.Column(0, iRow), 0) & ""), 0)
                End If
            End If
        Next
    End If
    Me.DrawnUnitCost = TotalCost
    Me.DrawSelectedCost = SelectedCost
    Me.DrawnSelectedCostWParts = SelectedCost + DrawnCostWParts

    TotalCost = 0
    SelectedCost = 0
    If OpenUnitsList.ListCount > 0 Then
        For iRow = 0 To (OpenUnitsList.ListCount)
            If Nz(OpenUnitsList.Column(4, iRow), 0) <> "$" Then
                TotalCost = TotalCost + Nz(OpenUnitsList.Column(4, iRow), 0)
            End If
            If OpenUnitsList.Selected(iRow) = True And Nz(OpenUnitsList.Column(4, iRow), 0) <> "$" Then
                SelectedCost = SelectedCost + Nz(OpenUnitsList.Column(4, iRow), 0)
                SelectedCostWParts = SelectedCostWParts + Nz(OpenUnitsList.Column(5, iRow), 0)
            End If
        Next
    End If
    Me.AvailUnitCost = TotalCost
    Me.OpenSelectedCost = SelectedCost
    OpenSelectedCostWParts = SelectedCostWParts
    'Remarked 8/18/2019
    TotalCost = 0
    If DrawnPartsList.ListCount > 0 Then
        For iRow = 1 To (DrawnPartsList.ListCount)
            TotalCost = TotalCost + Nz(DrawnPartsList.Column(6, iRow), 0)
        Next
    End If
    Me.ChildPartsCost = TotalCost
End Sub
Private Sub GetUnitCounts()
    Dim LabelOpen, LabelOpenParts, LabelDrawn, LabelDrawnParts As String
    'DrawnPartsList
    OpenUnitsX = 0
    DrawnUnitsX = 0
    OpenPartsX = 0
    DrawnPartsX = 0
    OpenUnitsX = [OpenUnitsList].[ItemsSelected].[Count]
    DrawnUnitsX = [DrawnUnitsList].[ItemsSelected].[Count]
    OpenPartsX = [OpenPartsList].[ItemsSelected].[Count]
    DrawnPartsX = [DrawnPartsList].[ItemsSelected].[Count]

    If OpenUnitsX = 1 Then LabelOpen = OpenUnitsX & " Unit" Else LabelOpen = OpenUnitsX & " Units"
    If OpenPartsX = 1 Then LabelOpenParts = OpenPartsX & " Part" Else LabelOpenParts = OpenPartsX & " Parts"

    If DrawnUnitsX = 1 Then LabelDrawn = DrawnUnitsX & " Unit" Else LabelDrawn = DrawnUnitsX & " Units"
    If DrawnPartsX = 1 Then LabelDrawnParts = DrawnPartsX & " Part" Else LabelDrawnParts = DrawnPartsX & " Units"

    Me.OpenSelectedCount = LabelOpen & " Selected"
    Me.OpenSelectedPartsCount = LabelOpenParts & " Selected"
    Me.AddUnits.caption = ChrW(10010) & " Draw " & LabelOpen & " "
    Me.AddParts.caption = ChrW(10010) & " Draw " & LabelOpenParts & " "
    Me.DrawSelectedCount = LabelDrawn & " Selected"
    Me.DrawnSelectedPartsCount = LabelDrawnParts & " Selected"
    Me.RemoveUnits.caption = ChrW(9940) & " Remove " & LabelDrawn & " "
    Me.RemoveParts.caption = ChrW(9940) & " Remove " & LabelDrawnParts & " "

    Call RunningTotal
End Sub
Public Sub ChildPartTypeList_Click()
    If Nz(Me.GroupID, 0) > 0 Then
        SQLz = "SELECT * FROM DrawDowns_ChildParts_Open " & _
            "WHERE GroupID=" & Me.GroupID & " " & _
            "AND VendorId= " & Me.ChildPartTypeList & " " & _
            "AND ChildType='" & Me.ChildPartTypeList.Column(5) & "' " & _
            "Order By CInvoice,AssetID;"
        Me.OpenPartsList.RowSource = SQLz
        SQLz = "SELECT * FROM DrawDowns_ChildParts_Drawn " & _
            "WHERE GroupID=" & Me.GroupID & " " & _
            "AND VendorId= " & Me.ChildPartTypeList & " " & _
            "AND ChildType='" & Me.ChildPartTypeList.Column(5) & "' " & _
            "Order By CInvoice,AssetID;"
        Me.DrawnPartsList.RowSource = SQLz
    End If
End Sub

Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*3" And Me(Ccrl.Name).Name Like "*3" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*4" And Me(Ccrl.Name).Name Like "*4" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    'If Me.PickModeToggle = True Then Call ReselectPickedUnits
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
    Call GetUnitCounts
End Function

Private Sub SyncChildInvoiceDrawn_Click()
    If [DrawnUnitsList].[ItemsSelected].[Count] < 1 Then
        Call MsgBox("You MUST Select Unit(s) First", vbCritical, "No Units Selected")
        Exit Sub
    End If
    DoCmd.SetWarnings False
    Dim InvDatex, InvPaidDatex As String
    Dim DrawnDowni As Long
    For Each ItemVar In Me.DrawnUnitsList.ItemsSelected
        InvX = DLookup("UnitInvoiceNum", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar))
        InvDatex = Nz(DLookup("UnitInvoiceDate", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar)), "")
        InvPaidDatex = Nz(DLookup("UnitVendorPytDate", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar)), "")
        DrawnDowni = Nz(DLookup("DrawDownId", "Units", "UnitId=" & Me.DrawnUnitsList.ItemData(ItemVar)), 0)

        SQLz = "Update ChildParts Set ChildInvoiceNum='" & InvX & "' "
        If Nz(InvDatex, "") <> "" Then SQLz = SQLz & ", ChildInvdt='" & InvDatex & "' "
        If Nz(InvPaidDatex, "") <> "" Then SQLz = SQLz & ", ChildVendorPytDate='" & InvPaidDatex & "' "
        If Nz(DrawnDowni, 0) <> 0 Then SQLz = SQLz & ", DrawDownID=" & DrawnDowni & " "
        SQLz = SQLz & "Where UnitREF=" & Me.DrawnUnitsList.ItemData(ItemVar) & ";"
        If InvX & "test" <> "test" Then DoCmd.RunSQL (SQLz)

    Next ItemVar
    If InvX & "test" <> "test" Then Call MsgBox("ChildParts Synced with Unit Draw, Invoice Num, and Dates", vbInformation, "Sync'd Parts")
    Me.DrawnPartsList.Requery
    DoCmd.SetWarnings True
End Sub

Private Sub SyncChildInvoiceOpen_Click()
    If [OpenUnitsList].[ItemsSelected].[Count] < 1 Then
        Call MsgBox("You MUST Select Unit(s) First", vbCritical, "No Units Selected")
        Exit Sub
    End If
    DoCmd.SetWarnings False
    Dim InvDatex, InvPaidDatex As String
    For Each ItemVar In Me.OpenUnitsList.ItemsSelected
        InvX = DLookup("UnitInvoiceNum", "Units", "UnitId=" & Me.OpenUnitsList.ItemData(ItemVar))
        InvDatex = Nz(DLookup("UnitInvoiceDate", "Units", "UnitId=" & Me.OpenUnitsList.ItemData(ItemVar)), "")
        InvPaidDatex = Nz(DLookup("UnitVendorPytDate", "Units", "UnitId=" & Me.OpenUnitsList.ItemData(ItemVar)), "")
        SQLz = "Update ChildParts Set ChildInvoiceNum='" & InvX & "' "
        If Nz(InvDatex, "") <> "" Then SQLz = SQLz & ", ChildInvdt='" & InvDatex & "' "
        If Nz(InvPaidDatex, "") <> "" Then SQLz = SQLz & ", ChildVendorPytDate='" & InvPaidDatex & "' "
        SQLz = SQLz & "Where UnitREF=" & Me.OpenUnitsList.ItemData(ItemVar) & ";"
        If InvX & "test" <> "test" Then DoCmd.RunSQL (SQLz)
    Next ItemVar
    If InvX & "test" <> "test" Then Call MsgBox("ChildParts Synced with Unit Invoice Num and Dates", vbInformation, "Sync'd Parts")
    DoCmd.SetWarnings True
End Sub




Private Sub ToggleSortAsset1_Click()
    Call SortListBox("ToggleSortAsset1", "OpenUnitsList", "AssetID")
End Sub
Private Sub ToggleSortAsset2_Click()
    Call SortListBox("ToggleSortAsset2", "DrawnUnitsList", "AssetID")
End Sub
Private Sub ToggleSortAsset3_Click()
    Call SortListBox("ToggleSortAsset3", "OpenPartsList", "AssetID")
End Sub
Private Sub ToggleSortAsset4_Click()
    Call SortListBox("ToggleSortAsset4", "DrawnPartsList", "AssetID")
End Sub

Private Sub ToggleSortInvoice1_Click()
    Call SortListBox("ToggleSortInvoice1", "OpenUnitsList", "Invoice")
End Sub
Private Sub ToggleSortInvoice2_Click()
    Call SortListBox("ToggleSortInvoice2", "DrawnUnitsList", "Invoice")
End Sub
Private Sub ToggleSortInvoice3_Click()
    Call SortListBox("ToggleSortInvoice3", "OpenPartsList", "CInvoice")
End Sub
Private Sub ToggleSortInvoice4_Click()
    Call SortListBox("ToggleSortInvoice4", "DrawnPartsList", "CInvoice")
End Sub

Private Sub ToggleSortOEC1_Click()
    Call SortListBox("ToggleSortOEC1", "OpenUnitsList", "Format(IFAmount,'$#,#')")
End Sub
Private Sub ToggleSortOEC2_Click()
    Call SortListBox("ToggleSortOEC2", "DrawnUnitsList", "Format(OEC,'$#,#')")
End Sub
Private Sub ToggleSortCost3_Click()
    Call SortListBox("ToggleSortOEC2", "OpenPartsList", "PCost")
End Sub
Private Sub ToggleSortCost4_Click()
    Call SortListBox("ToggleSortOEC2", "DrawnPartsList", "CCost")
End Sub

Private Sub ToggleSortPartType3_Click()
    Call SortListBox("ToggleSortPartType3", "OpenPartsList", "ChildType")
End Sub
Private Sub ToggleSortPartType4_Click()
   Call SortListBox("ToggleSortPartType4", "DrawnPartsList", "ChildType")
End Sub

Private Sub ToggleSortUnitNum1_Click()
    Call SortListBox("ToggleSortUnitNum1", "OpenUnitsList", "UnitNum")
End Sub
Private Sub ToggleSortUnitNum2_Click()
    Call SortListBox("ToggleSortUnitNum2", "DrawnUnitsList", "UnitNum")
End Sub
Private Sub ToggleSortUnitNum3_Click()
    Call SortListBox("ToggleSortUnitNum3", "OpenPartsList", "UnitNum")
End Sub
Private Sub ToggleSortUnitNum4_Click()
    Call SortListBox("ToggleSortUnitNum4", "DrawnPartsList", "UnitNum")
End Sub

Private Sub ToggleSortVIN1_Click()
    Call SortListBox("ToggleSortVIN1", "OpenUnitsList", "VIN")
End Sub
Private Sub ToggleSortVIN2_Click()
    Call SortListBox("ToggleSortVIN2", "DrawnUnitsList", "VIN")
End Sub
Private Sub ToggleSortVIN3_Click()
    Call SortListBox("ToggleSortVIN3", "OpenPartsList", "VIN")
End Sub
Private Sub ToggleSortVIN4_Click()
    Call SortListBox("ToggleSortVIN4", "DrawnPartsList", "VIN")
End Sub

Private Sub ToggleSortVendorPaid2_Click()
    Call SortListBox("ToggleVendorPaid2", "DrawnUnitsList", "UnitVendorPytDate")
End Sub

Private Sub TransactionReport_Click()
    If Me.DrawID & "test" = "test" Then
        MsgBox "Please Select a Draw", vbCritical, "Need A Draw"
        Exit Sub
    End If
    DoCmd.OpenReport "DrawDown_TransactionSummary_Report", acViewPreview

End Sub
