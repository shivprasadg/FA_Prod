Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub AddDocLink_Click()
    DoCmd.close acForm, "DocAddEditLinkForm"
    DoCmd.OpenForm "DocAddEditLinkForm", acNormal, , , acFormAdd, , Me.clientID & "|" & Me.MLID & "|" & Me.SchID & "|0|0|" & Me.Form.Name & "|Add"
End Sub

Private Sub Form_Current()
    SQLz = "SELECT Deals.faid, Deals.dealShortDesc From Deals WHERE ClientRef=" & Me.clientID & " ORDER BY Deals.faid DESC;"
    Me.SchFAID.RowSource = SQLz
    Me.GroupFundList.Requery
    Me.BillingRules.Visible = False
    If FAUserAccessLevel = 99 And Me.MLOrig <> "FA" Then Me.BillingRules.Visible = True

    SQLz = "SELECT contactID, FullName, Title, ClientID FROM ClientContacts_Base WHERE ClientID= " & Me.clientID & " Order by FullName;"
    Me.SchSigner.RowSource = SQLz
    Me.SchSigner2.RowSource = SQLz
    Me.SchSigner3.RowSource = SQLz
    'Me.RALSignerID.RowSource = SQLz
    '   Me.EventHistory.Requery
End Sub
Private Sub AddFHUT_Click()
    On Error GoTo FixMe
    If DCount("SchID", "FHUT", "isnull(FHUTPD) AND SchID=" & Me.SchID) > 0 Then
        MsgBox "There is an exisiting non-paid FHUT for this Schedule of units" & vbNewLine & "The FHUT Form will open with the exisiting FHUT", vbCritical, "No need to add"
        DoCmd.OpenForm "FHUTAddUpdate", acNormal, , , acFormEdit
        Exit Sub
    End If
    If DLookup("MLOrig", "vw_SixKeys", "SchID=" & Me.SchID) <> "FA" Then
        MsgBox "Only FA leased units can have an FHUT submission." & vbNewLine & "The FHUT Form will not open", vbCritical, "Not FA Lease"
        Exit Sub
    Else
        Dim Committ As String
        Committ = MsgBox("Are you sure you want to add a FHUT submission for this Schedule of Units?", vbQuestion + vbYesNo, "Please confrim")
    End If
    If Committ = vbNo Then Exit Sub
    SQLz = "INSERT INTO fhut ( FHUTURef, GroupID, SchID ) " & _
           "SELECT AssetID, SGrpID, SchID FROM vw_SixKeys Where Schid=" & Me.SchID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    DoCmd.OpenForm "FHUTAddUpdate", acNormal, , , acFormEdit
    Exit Sub
FixMe:
    MsgBox Err.Description, , "Update Stopped"
End Sub

Private Sub AddGroup_Click()
    DoCmd.close acForm, "GroupAdd_Form"
    DoCmd.OpenForm "GroupAdd_Form", , , , , , Me.SchID & "|" & Me.SchFAID
    Forms!GroupForm_Add.SchID = Me.SchID
End Sub

Private Sub CloseMe_Click()

    If IsNull(Me.SchFAID) Then
        Select Case MsgBox("Do you want to save the new Schedule?", vbYesNo)
        Case vbYes
            If IsNull(Me.SchFAID) Then
                MsgBox "Please update the FAID and any other required fields"
                Exit Sub
            End If
        Case vbNo
            'MsgBox "No Changes saved to Unit"
            Me.Undo
            Me.Undo
        End Select
        Exit Sub
    End If
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.DocLinksSub.Form.RecordSource = "Select * From DocsLinkSubQ WHERE SchID=" & Me.SchID
End Sub

Private Sub GroupFundList_DblClick(Cancel As Integer)
    If Me.GroupFundList.Column(4) & "test" <> "test" Then
        DoCmd.close acForm, "GroupForm"
        DoCmd.OpenForm "GroupForm", , , "sGrpID=" & Me.GroupFundList.Column(4)
        Forms!GroupForm.Tabcontrol.Value = 5
    End If
End Sub

Private Sub SchBnkClNO_AfterUpdate()
    Me.SchBnkClNO = Trim(Me.SchBnkClNO)
    If Right(Me.SchBnkClNO, 1) = "-" Or Me.SchBnkClNO Like "*--*" Then
        MsgBox "You cannot use dashes like that...", vbCritical, "Wrong Format"
        Me.SchBnkClNO = Replace(Me.SchBnkClNO, "-", "")
        If Len(Me.SchBnkClNO) > 3 Then Me.SchBnkClNO = Left(Me.SchBnkClNO, 3) & "-" & Mid(Me.SchBnkClNO, 4, Len(Me.SchBnkClNO) - 3)
    End If
End Sub

Private Sub SchSigner_Click()
    Me.SchSignerID = Me.SchSigner.Column(0)
End Sub
Private Sub SchSigner2_Click()
    Me.SchSigner2ID = Me.SchSigner2.Column(0)
End Sub
Private Sub SchSigner3_Click()
    Me.SchSigner3ID = Me.SchSigner3.Column(0)
End Sub

Private Sub UnitSummary_Click()
    Call MakeScheduleUnitSummary(Me.clientID, Me.SchID, Me.SchNo)
End Sub

Private Sub OpenClientsForm_Click()
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
End Sub

Private Sub RentReport_Click()
    Dim XLS As String
    Dim Countx As Integer
    XLS = "|E,BC|K,BC|L,TC|I,BC|M,P1C|N,$E|O,$E|P,$E|Q,$E|"
    If IsNull(Me.Location) Then
        Countx = DCount("UnitVin", "Schedule_Rent_Report")
        FormatExcel "Schedule_Rent_Report", "C:\Faas\" & FAUserName & "\Temp\Schedule_Rent_Report.xlsx", "Schedule Rent", , "" & XLS & "", Countx, Me.SchDesc & " - Rent Report - " & Countx & " Units", True
    Else
        Countx = DCount("UnitVin", "Location_Rent_Report")
        FormatExcel "Location_Rent_Report", "C:\Faas\" & FAUserName & "\Temp\Location_Rent_Report.xlsx", "Location Rent", , "" & XLS & "", Countx, "Location: " & Me.Location.Column(1) & " - Rent Report - " & Countx & " Units", True
    End If
End Sub

Private Sub SchBillAdd_Change()
    MsgBox "This Changed The 'Bill To' Address For ALL Groups In This Schedule", vbInformation, "Cascade Change"
    SQLz = "UPDATE SchGrp SET SchGrp.BillToAddressID =" & Me.SchBillAdd & " Where SGrpSchID =" & Me.SchID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub

Private Sub UPDtFHUT_Click()
    DoCmd.OpenForm "Update FHUT"
End Sub

Private Sub RALExport_Click()
    Dim XLS As String
    Dim Countx As Integer
    'On Error GoTo fixme
    SQLz = "Select * From RentAssignmentLetterExport Where Schid=" & Me.SchID & ";"
    Countx = DCount("VIN", "RentAssignmentLetterExport", "Schid=" & Me.SchID)
    XLS = "|A,BL|B,BC|D,BC|H,BR|I,$E|J,BC|L,BC|N,HH|O,HH|"
    SendReport2Excel SQLz, "RAL", "Rent Assigment Letter Export For: " & [ClientShNm] & " Schedule " & [SchNo] & " Units: " & Countx, True, XLS, Countx
FixMe:
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

    Select Case DataErr
    Case 3146    'Required Field ODBC Error 'Case 3146 To 3621 'ODBC Error
        MsgBox "You must have data in all Yellow Required Fields before moving off the record or closing the form.", vbCritical, "Required Field Error #" + str(DataErr) + " - All Required Fields Must Contain Data!"
        Response = acDataErrContinue
    Case 3146 To 3621
        MsgBox "Unknown ODBC error occurred", vbCritical, "ODBC Error #" & str(DataErr)
        'Response = acDataErrContinue
    Case Else
        MsgBox AccessError(DataErr)
        'Response = acDataErrContinue
    End Select
End Sub

Private Sub ViewGrpBtn_Click()
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", acNormal, , "[SchID]=" & SchID, , , DMin("sGrpID", "SchGrp", "sGrpSchId =" & Me.SchID)
End Sub
