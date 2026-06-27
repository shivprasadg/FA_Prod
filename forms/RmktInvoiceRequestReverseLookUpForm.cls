Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'Dim SQLz       As String
'
'Private Sub Form_Load()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'End Sub
'Private Sub AssetSearch_Change()
''Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where AssetID Like'*" & Me.AssetSearch.Text & "*' Order By AssetID;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'
'Private Sub BuyerSearch_Change()
'    Me.AssetSearch = ""
'    'Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where Buyer Like'*" & Me.BuyerSearch.Text & "*' Order By AssetID;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'
'Private Sub DescSearch_Change()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    'Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where UnitDesc Like'*" & Me.DescSearch.Text & "*' Order By UnitDesc;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'
'Private Sub InvNumSearch_Change()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    'Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where InvNum Like'*" & Me.InvNumSearch.Text & "*' Order By AssetID;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'Private Sub OpenInvoice_Click()
'    If FAUserAccess > 11 Or FAUserAccessGroup Like "*Acct*" Then
'        If IsNull(Me.InvoicedUnitsList.Column(0)) Then Exit Sub
'        DoCmd.OpenForm "ReMarketInvoiceManager"
'        Forms!ReMarketInvoiceManager.Form.BuyerRefCBO = Me.InvoicedUnitsList.Column(0)
'        Form_ReMarketInvoiceManager.RemoteBuyerRef
'        Forms!ReMarketInvoiceManager.Form.InvoiceList = Me.InvoicedUnitsList.Column(2)
'        Form_ReMarketInvoiceManager.RemoteInvoice
'    Else
'        MsgBox "Sorry your don't have access Rights to that form", , "ATLAAS Security"
'    End If
'End Sub
'Private Sub OpenBuyer_Click()
'    If IsNull(Me.InvoicedUnitsList.Column(0)) Then Exit Sub
'    DoCmd.OpenForm "BuyersForm"
'    Forms!BuyersForm.BuyerList = Me.InvoicedUnitsList.Column(0)
'    Form_BuyersForm.BuyerList_Click
'    Forms!BuyersForm.RemarketingInvoiceRequest_Sub.Form.InvoiceRequestID = Me.InvoicedUnitsList.Column(2)
'    Forms!BuyersForm.TabControls.Value = 4
'    Pause (0.25)
'    Form_BuyersForm.RemarketingInvoiceRequest_Sub.Form.InvoiceRequestID_Click
'     Pause (0.25)
'    Forms!BuyersForm.RemarketingInvoiceRequest_Sub.Form.AssignedUnitsList = Me.InvoicedUnitsList.Column(5)
'End Sub
'
'Private Sub OpenSchedule_Click()
'        If IsNull(Me.InvoicedUnitsList.Column(0)) Then Exit Sub
'        DoCmd.OpenForm "ScheduleControl"
'        Forms!ScheduleControl.Form.SearchAssetID2.SetFocus
'        Forms!ScheduleControl.Form.SearchAssetID2.Text = Me.InvoicedUnitsList.Column(5)
'        Forms!ScheduleControl.Form.UnitLookUp = Me.InvoicedUnitsList.Column(5)
'        Form_ScheduleControl.Remotelookup
'
'End Sub
'
'Private Sub OpenUnit_Click()
'    If IsNull(Me.InvoicedUnitsList.Column(0)) Then Exit Sub
'    Call OpenUnitsForm(Me.InvoicedUnitsList.Column(5))
'    'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.InvoicedUnitsList.Column(5)
'End Sub
'
'
'Private Sub ReqSearch_Change()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    'Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where Req Like'*" & Me.ReqSearch.Text & "*' Order By Req;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'
'Private Sub ScheduleCheckLists_Click()
'    If CheckMe = False Then Exit Sub
'    Call OpenBuyersFrom(Me.InvoicedUnitsList.Column(0))
'End Sub
'
'Private Sub UnitNumSearch_Change()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    'Me.UnitNumSearch = ""
'    Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where UnitNum Like'*" & Me.UnitNumSearch.Text & "*' Order By UnitNum;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'
'Private Sub VINSearch_Change()
'    Me.AssetSearch = ""
'    Me.BuyerSearch = ""
'    Me.DescSearch = ""
'    Me.InvNumSearch = ""
'    Me.ReqSearch = ""
'    Me.UnitNumSearch = ""
'    'Me.VINSearch = ""
'    SQLz = "Select * From RmktInvoiceRequestReverseLookUp Where VIN Like'*" & Me.VINSearch.Text & "*' Order By Vin;"
'    Me.InvoicedUnitsList.RowSource = SQLz
'End Sub
'Public Function CheckMe() As Boolean
'    If Me.InvoicedUnitsList.[ItemsSelected].[Count] < 1 Then
'        MsgBox "You Must Select A Unit", vbInformation, "No Unit"
'        CheckMe = False
'    Else
'        CheckMe = True
'        Exit Function
'    End If
'
'End Function
'
