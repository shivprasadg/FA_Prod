Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String

Private Sub AddNewCompany_Click()
    DoCmd.OpenForm "AddClientGroup"
End Sub

Private Sub Client_Lease_Summary_Click()
  Dim XLS As String
    Dim Countx As Integer

    Countx = DSum("Units", "Client_Lease_Summary", "ClientGroupID=" & Me.ClientGroupID)

    If Nz(Countx, 0) <= 0 Then
        MsgBox "No Matching Units", vbInformation, "No Result"
        Exit Sub
    End If

    MessageUser "Open", "Making Detailed Tab...", "This will take a moment or two..."
    Dim CGroupX As String
    SQLz = "Select * From Client_Lease_Detail Where ClientGroupID=" & Me.ClientGroupID & " Order By Client, SchID;"

    XLS = ""
    XLS = "|A,BL|I,$E|J,$E|K,$E|O,BR|R,BC|S,$R|T,$E|V,$E|W,$E|X,P4|Y,BC|Z,BC|AA,BC|AB,$E|AD,ZZ|AD,ZZ|AD,ZZ|AD,BC|"
    SendReport2Excel SQLz, "Unit_Detail", "Lease Unit Detail Report for Client Group: " & CGroupX & " All Schedules - Units: " & Countx, True, XLS, 0, False, 1, False

    MessageUser "Update", "Making Summary Tab...", "Excel will open in a moment..."
    SQLz = "Select * From Client_Lease_Summary Where ClientGroupID=" & Me.ClientGroupID & " Order By Client, SchID;"
    CGroupX = DLookup("ClientGroupName", "ClientGroups", "ClientGroupId=" & Me.ClientGroupID)
    XLS = "|A,BL|G,TC|G,CE|J,P5|K,$E|L,HH|M,HH|"
    SendReport2Excel SQLz, "Summary", "Lease Summary Report for Client Group: " & CGroupX & " All Schedules Summary - Units: " & Countx, True, XLS, 0, True, 2, True

    MessageUser "Close"
End Sub

Private Sub clientCompanyName_DblClick(Cancel As Integer)
    Call LockIt
End Sub
Private Sub LockIt()
    If Me.LockUnLock.caption = "Locked" Then
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Required" Then
                Me(Ccrl.Name).Locked = False
            End If
        Next Ccrl
        Me.LockUnLock.caption = "UnLocked"
    Else
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Required" Then
                Me(Ccrl.Name).Locked = True
            End If
        Next Ccrl
        Me.LockUnLock.caption = "Locked"
    End If
End Sub

Private Sub FleetListProfile_Click()
    Dim XLS, Headerx, SQLz As String

    SQLz = SQLz & "Select Client"            '- Col:A
    SQLz = SQLz & " ,Sch"                    '- Col:B
    SQLz = SQLz & " ,Group"                  '- Col:C
    SQLz = SQLz & " ,AssetID"                '- Col:D
    SQLz = SQLz & " ,UnitNum"                '- Col:E
    SQLz = SQLz & " ,UnitVIN"                '- Col:F
    SQLz = SQLz & " ,Status"                 '- Col:G
    SQLz = SQLz & " ,InService"              '- Col:H
    SQLz = SQLz & " ,Garage"                 '- Col:I
    SQLz = SQLz & " ,State"                  '- Col:J
    SQLz = SQLz & " ,UnitYr"                 '- Col:K
    SQLz = SQLz & " ,UnitMake"               '- Col:L
    SQLz = SQLz & " ,UnitModel"              '- Col:M
    SQLz = SQLz & " ,UnitType"               '- Col:N
    SQLz = SQLz & " ,UnitSubType"            '- Col:O
    SQLz = SQLz & " ,Axle"                   '- Col:P
    SQLz = SQLz & " ,UnitConfig"             '- Col:Q
    SQLz = SQLz & " ,Engine"                 '- Col:R
    SQLz = SQLz & " ,Trans "                 '- Col:S
    SQLz = SQLz & " ,OEC"                    '- Col:T
    SQLz = SQLz & " ,LeaseType"              '- Col:U
    SQLz = SQLz & " ,LeaseTerm"              '- Col:V
    SQLz = SQLz & " ,MilesAllowed"           '- Col:W
    SQLz = SQLz & " ,MileCharge"             '- Col:X
    SQLz = SQLz & " ,BLCD"                   '- Col:Y
    SQLz = SQLz & " ,LeaseExp"               '- Col:Z
    SQLz = SQLz & " ,EXchangedBy"            '- Col:AA
    SQLz = SQLz & " ,EXchangedByUnit"        '- Col:AB
    SQLz = SQLz & " ,EXchangedFor"           '- Col:AC
    SQLz = SQLz & " ,EOptionSelected"        '- Col:AD
    SQLz = SQLz & " ,ODOSource"              '- Col:AE
    SQLz = SQLz & " ,ODORead"                '- Col:AF
    SQLz = SQLz & " ,ODODate"                '- Col:AG
    SQLz = SQLz & " ,OverMiles"              '- Col:AH
    SQLz = SQLz & " ,CurrRent"               '- Col:AI
    SQLz = SQLz & " ,CurrRentType"           '- Col:AJ
    SQLz = SQLz & " ,Extends"                '- Col:AK
    SQLz = SQLz & " ,ClientGroupID "         '- Col:AL
    SQLz = SQLz & " FROM vw_ExitReport0 WHERE ClientGroupID=" & Me.ClientGroupID & ";"

    XLS = "|A,DD3|A,TR|A,W15|"
    XLS = XLS & "B,TC|B,W07|"
    XLS = XLS & "C,TC|C,W07|"
    XLS = XLS & "D,TX|D,BC|D,W08|"
    XLS = XLS & "E,TC|E,W08|"
    XLS = XLS & "F,TL|F,W20|"
    XLS = XLS & "G,BC|G,cLabel:End=G:3:Unit Status|G,cColor:3:0.0.0|G,cBorder:Med:End=G|G,W06|"
    XLS = XLS & "H,TC|H,W11|"
    XLS = XLS & "I,W25|I,TL|"
    XLS = XLS & "J,TC|J,W05|"
    XLS = XLS & "K,TC|K,cLabel:End=K:3:MYear|K,W06|"
    XLS = XLS & "L,TL|L,cLabel:End=L:3:Make|L,W14|"
    XLS = XLS & "M,TL|M,cLabel:End=M:3:Model|M,W13|"
    XLS = XLS & "N,TL|N,cLabel:End=N:3:UType|N,W11|"
    XLS = XLS & "O,TL|O,cLabel:End=O:3:SubType|O,W12|"
    XLS = XLS & "P,TC|P,cLabel:End=P:3:Axle (type)|P,W06|"
    XLS = XLS & "Q,TC|Q,cLabel:End=Q:3:Config (UOM)|Q,W08|"
    XLS = XLS & "R,TC|R,W15|"
    XLS = XLS & "S,TC|S,cLabel:End=S:3:Trans|S,W16|"
    XLS = XLS & "T,$E|T,cLabel:End=T:3:Total OEC (w parts)|T,W14|"
    XLS = XLS & "U,TC|U,cLabel:End=U:3:Lease Type|U,W09|"
    XLS = XLS & "V,BC|V,cLabel:End=V:3:Lease Term|V,W06|"
    XLS = XLS & "W,SR|W,cLabel:End=W:3:Miles Allowed|W,cLabel:End=W:2:" & TLeftLine & LineSpace(1) & "Allowed M/H Per Year " & LineSpace(10) & TRightLine & "|W,cColor:2:0.0.128:End=X|W,cFontColor:2:255.255.255:End=X|W,W10|"
    XLS = XLS & "X,$R|X,TC|X,cLabel:End=X:3:Mile Charge|X,W06|"
    XLS = XLS & "Y,W11|"
    XLS = XLS & "Z,cLabel:End=Z:3:Lease Expire|Z,W11|"
    XLS = XLS & "AA,TC|AA,cLabel:End=AA:3:Exchanged Sch-Grp|AA,cColor:3:59.112.6|AA,W12|"
    XLS = XLS & "AB,TC|AB,cLabel:End=AB:3:Exch'd Unit|AB,cColor:3:59.112.6|AB,W08|"
    XLS = XLS & "AC,TC|AC,cLabel:End=AC:3:Replaced Sch-Grp|AC,cColor:3:47.117.181|AC,W12|"
    XLS = XLS & "AD,TC|AD,cLabel:End=AD:3:Exit Option|AD,cColor:3:51.51.0|AD,W16|"
    XLS = XLS & "AE,TC|AE,cLabel:End=AE:3:ODO Source|AE,W06|"
    XLS = XLS & "AF,Sint|AF,TC|AF,cLabel:End=AF:3:ODO Read|AF,W10|"
    XLS = XLS & "AG,TC|AG,cLabel:End=AG:3:ODO Date|AG,W11|"
    XLS = XLS & "AH,Sint|AH,cLabel:End=AH:3:Over Miles Driven|AH,cColor:3:191.143.0|AH,cBorder:Med:End=AH|AH,W11|"
    XLS = XLS & "AI,$E|AI,cLabel:End=AI:3:Curr Rent|AI,cColor:3:191.143.0|AI,W13|"
    XLS = XLS & "AJ,TC|AJ,cLabel:End=AJ:3:Rent Type|AJ,cColor:3:191.143.0|AJ,W07|"
    XLS = XLS & "AK,W15|"
    XLS = XLS & "AL,HH|"

    Headerx = "Fleet Profile Report: " & Me.ClientShNm & " : " & Date
    DoCmd.SetWarnings True
    MessageUser "Open", "Excel - Fleet Profile  Report", "Please Wait this will take a moment..."
    SendReport2Excel SQLz, "FleetProfile", Headerx, True, XLS, 0, False
    MessageUser "Close"
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_ClientsView"
End Sub

Private Sub GroupFleetList_Click()
    On Error GoTo GroupFleetList_Click_Error
    Call ExportFleetList(Me.ClientGroupID.Column(1), True, Me.clientID, Me.ClientGroupID)
    On Error GoTo 0
    Exit Sub
GroupFleetList_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GroupFleetList_Click of VBA Document Form_ClientsView"
End Sub
Private Sub ClientFleetList_Click()
    On Error GoTo ClientFleetList_Click_Error
    Call ExportFleetList(Me.ClientShNm, False, Me.clientID, Me.ClientGroupID)
    On Error GoTo 0
    Exit Sub
ClientFleetList_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClientFleetList_Click of VBA Document Form_ClientsView"
End Sub

Private Sub ClientList_Click()
    Me.Form.RecordSource = "Select * From Clients Where ClientID=" & Me.ClientList.Column(1) & ""
    Me.Form.Filter = "clientID=" & Me.ClientList.Column(1) & ""
    Me.Form.FilterOn = True
    Me.ClientFleetList.caption = SplitString(Me.clientCompanyName, " ", 0) & "'s Fleet List"
    Me.GroupFleetList.caption = "Fleets in " & Me.ClientGroupID.Column(1) & "'s Group"

    Forms!ClientsView!Locations_sub.Form.SetLocations
    Forms!ClientsView!proposalsSummary_sub.Form.SetDeals
    Forms!ClientsView!ClientContact_sub.Form.SetContacts
    Forms!ClientsView!SUF_Conditions.Form.SetSuf
    Forms!ClientsView!ValueAdded_sub.Form.SetValueAdd
    Forms!ClientsView!Leasing_sub.Form.SetLease
End Sub

Private Sub ClientShNm_DblClick(Cancel As Integer)
    Call LockIt
End Sub

Private Sub clientType_DblClick(Cancel As Integer)
    Call LockIt
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ImportFAID_Click()
If FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Or FAUserAccessGroup Like "*SYND*" Then
DoCmd.OpenForm "ATF_SFImporter"
Forms!ATF_SFImporter.SFAccountID = Me.SFAccountID
Forms!ATF_SFImporter.FAIDLPAFList = "0"
Else
MsgBox "Sorry you do not have access to the Salesforce Importer", vbCritical, AtlaasVersion
End If
End Sub

Private Sub LOinitial_Click()
    If IsNull(Me.OrigBDEX) Then Me.OrigBDEX = Me.LOinitial.Column(0)
    Me.SalespersonREF = Me.LOinitial.Column(2)
End Sub

Private Sub LOinitial_DblClick(Cancel As Integer)
    Call LockIt
End Sub


Private Sub CloseformBtn_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Transactions_Click()
    DoCmd.Requery "Forms!ClientsView.Deals_sub"
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT Client, clientID FROM ClientsList " & _
           "WHERE Client Like '*" & Me.SearchClient.Text & "*' Or ClientShNm Like '*" & Me.SearchClient.Text & "*' ORDER BY Client;"
    Me.ClientList.RowSource = SQLz
End Sub


Private Sub Form_Load()
    Me.ClientList.RowSource = "SELECT Client, clientID FROM ClientsList Order By Client;"
    If Nz(Me.OpenArgs, "x") <> "x" Then
        Me.ClientList = Me.OpenArgs
        Call ClientList_Click
    End If
    Me.SearchClient.SetFocus
End Sub
Private Sub Form_Current()

    On Error GoTo Form_Current_Error

    Me.LockUnLock.caption = UnLocked
    Call LockIt
    Me.ClientsView_Schedules1.Form.Refresh
    If Me.NewRecord Then
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Required" Then
                Me(Ccrl.Name).Locked = False
            End If
        Next Ccrl
        Me.LockUnLock.caption = "UnLocked"
    End If

    If FAUserAccessGroup Like "*TMT*" And FAUserAccessLevel > 50 Then
        Me.ImportFAID.Visible = True
        Me.ExtendedRentdiscount.Locked = False
    Else
        Me.ImportFAID.Visible = False
        Me.ExtendedRentdiscount.Locked = True
    End If
    On Error GoTo 0
    Exit Sub

Form_Current_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Current of VBA Document Form_ClientsView"
End Sub

Private Sub SFAccountID_DblClick(Cancel As Integer)
If FAUserAccess > 90 Then Me.SFAccountID.Locked = False
End Sub
