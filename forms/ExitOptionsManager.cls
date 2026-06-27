Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLx, SQLy, SQLz, SQLGrp, SQLBase, OnOff, OpenExitTxt, EndType, OrderByx As String
Dim Queryz As String
Dim GroupBY, WhereX As String
Dim SchIDz     As Integer

Private Sub SyncGetLatestData_Click()
Call MakeExitOptionLocal(True)
Call ResetForm_Click
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.AssetID
    Me.ScheduleIDx = Me.SchGrpList.Column(0)
    Me.GroupIDx = Me.groupID
    'dd = DLookup("MLNO", "vw_SixKeys", "SchID=" & Me.SchID)
    OpenExitTxt = "Open EXIT Options: "
    If Me.LeaseTypetoShow = 1 Then OpenExitTxt = "Open EXIT Options: " Else OpenExitTxt = "Select an Options"
    If Nz(Me.SchGrpList.Column(0), 0) <> 0 Then
        Me.OpenExitOptions.caption = OpenExitTxt & ClientList.Column(1) & " | " & Me.ScheduleGroup
        Else
        Me.OpenExitOptions.caption = OpenExitTxt
    End If
FixMe:
End Sub


Public Sub Form_Load()
    Call MakeExitOptionLocal
    Queryz = "ExitOptionsGroup_Local"
    Call ClearBoxes
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    Me.SearchClient = ""
    Me.SearchML = ""
    SQLGrp = ""
    EndType = "EOL"
    HavingX = "HAVING MLOrig='FA' "
    OpenExitTxt = "Open EXIT Options: "
    SQLGrp = SQLGrp & "SELECT GroupID, [Schedule] & ' | ' & [UnitGroup] AS [Sch|Grp], MLOrig, Count(AssetID) AS Units "
    SQLGrp = SQLGrp & "From " & Queryz & " "
    WhereX = "WHERE clientID = " & Nz(Me.ClientList.Column(0), 0) & " "
    GroupBY = "GROUP BY GroupID, MLOrig, Schedule, UnitGroup Order By GroupID Desc"
    Me.SchGrpList.RowSource = SQLGrp & WhereX & GroupBY & ";"
    Me.Move Me.WindowLeft, Height:=15000
End Sub

Public Sub LeaseTypeToShow_Click()
Exit Sub
'this routine will be removed - 3/31/2023
   '  If Me.LeaseTypetoShow = 1 Then
        WhereX = "WHERE MLOrig ='FA' "
          For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "FAOnly" Then
                Me(Ccrl.Name).Visible = True
            End If
        Next Ccrl
        Me.ModifyExch.caption = "EXCH"
        Me.LinkLineFA.Visible = True
       ' Me.LinkLinePort.Visible = False
        OpenExitTxt = "BuyIT Options: "
        Me.ToggleFindExchanged.caption = "Were Exchanged"
        Me.ToggleFindNOTExchanged.caption = "Yet To Exchange"
        Me.ExchGroupLabel.caption = "EXchangeIT Equipment"
        SQLz = "SELECT clientID, Client, Type FROM ClientMasterLeasesTypeCount_Summed Where Type = 'FA' or  Type = 'Both' Order by Client;"

    Me.ClientList.RowSource = SQLz
    Me.SearchClient = ""
    Me.SearchML = ""
    If Not IsNull(Me.ClientList.Column(0)) Then WhereX = WhereX & "AND ClientID=" & Me.ClientList.Column(0) & " "
    Me.SchGrpList.RowSource = SQLGrp & WhereX & GroupBY & ";"
    Me.OpenExitOptions.caption = OpenExitTxt & " Form"
    Me.Form.RecordSource = "Select Distinct * From ExitOptionsGroup_Local " & WhereX
End Sub

Private Sub SetToggles(Optional ExcludeToggle As String)
    Dim ctrl   As Control
    For Each ctrl In Me.Controls
        If ctrl.Name Like "*Toggle*" And ctrl.Name <> ExcludeToggle Then Me(ctrl.Name) = False
    Next
End Sub

Private Sub ResetForm_Click()
    Call SetToggles
    Toggle120Eol.caption = "Lease Exp < 180 Days"
    Toggle120Exch.caption = "ExchangeIT < 180 Days"
    Call Form_Load
    SQLz = "SELECT clientID, Client, Type FROM ClientMasterLeasesTypeCount_Summed;"
    Me.ClientList.RowSource = SQLz
    SQLz = "SELECT clientID, Client, Type FROM ClientMasterLeasesTypeCount_Summed Where Type = 'FA' or  Type = 'Both' Order by Client;"
    Me.ClientList.RowSource = SQLz
    Call ClearBoxes
    Call ClearList("SchGrpList")
    Call ClearList("ClientList")
    'Me.LeaseTypetoShow = 1
    Call LeaseTypeToShow_Click
    Me.OpenExitOptions.caption = "Open EXIT Options: "
    SQLz = "Select Distinct * From " & Queryz & " Where MLOrig='FA' Order By Client;"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub OpenExitOptions_Click()
'If Me.SchGrpList & "test" = "test" Then
'MsgBox "You did not select a Sch-Grp", vbCritical, "What's wrong?"
'Exit Sub
'End If
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    If Me.SchGrpList.Column(0) & "test" <> "test" Then
        SchIDz = Me.SchID
        Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
        Form_EXITOptions.ClientList_Click
        Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
        Form_EXITOptions.Leaselist_Click
        Forms!ExitOptions.ScheduleList = SchIDz
        Form_EXITOptions.ScheduleList_Click
        Forms!ExitOptions.GroupList = Me.GroupIDx
        Form_EXITOptions.GroupList_Click
    End If
End Sub

Private Sub ExitOptionsHistory_Click()
    If Me.SchGrpList.Column(0) & "test" = "test" Then
        MsgBox "You Must Select A MasterLease/Schedule/Group", , "No Group Selected"
        Exit Sub
    End If
    SchIDz = Me.SchID
        If DCount("AssetID", "EXITOptionsGroup_Local", "[Selected] <> '----  ----  ----' AND Schid=" & SchIDz) < 1 Then
        MsgBox "There is no history for this Schedule/Group" & vbNewLine & "The Exit Options Form will NOT open.", vbInformation, "No History"
        Exit Sub
    End If
    MessageUser "Open", "Getting History", "Opening Exit Options"
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = Me.groupID
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.History_Click
    MessageUser "Close"
End Sub

Public Sub ClientList_Click()
    Me.Detail.Visible = True
    Call ClearList("SchGrpList")

    SQLz = SQLBase & " AND MLOrig='FA' AND ClientId =" & Me.ClientList & " Order By Client, ReplacedSchNo;"
    Me.Form.RecordSource = SQLz
    Me.SearchML.Enabled = True
    Me.SchGrpList.Enabled = True
    Me.SearchML = ""
    SQLx = ""
    WhereX = " WHERE MLOrig='FA' AND clientID = " & Me.ClientList.Column(0) & " "
    Me.SchGrpList.RowSource = SQLGrp & WhereX & GroupBY & ";"
    SQLx = ""
    Call SetToggles
    Toggle120Eol.caption = "Lease Exp < 180 Days"
    Toggle120Exch.caption = "ExchangeIT < 180 Days"
    Me.UnitCount.Visible = True
End Sub

Private Sub ScheduleGroup_DblClick(Cancel As Integer)
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SchID=" & Me.SchID, , , Me.groupID
    Forms!GroupForm.Tabcontrol.Value = 8
End Sub

Private Sub NewSchGrp_DblClick(Cancel As Integer)
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SchID=" & Me.ReplacedByScheduleID, , , Me.ReplacedByGroupID
    Forms!GroupForm.Tabcontrol.Value = 8
End Sub

Public Sub SchGrpList_Click()
    SQLz = SQLBase & " AND GroupID =" & Me.SchGrpList & " Order By Client, SchID, AssetID ;"
    Me.Form.RecordSource = SQLz
    Me.ScheduleIDx = Me.SchGrpList.Column(0)
    Me.UnitCount.Visible = True
    Me.SearchClient = ""
End Sub

Private Sub ClearAll_Click()
    Me.SearchClient = ""
    Me.SearchML = ""
    Call ClearList("SchGrpList")
    Call ClearList("ClientList")
End Sub

Private Sub SearchClient_GotFocus()
    Me.SchGrpList.RowSource = ""
End Sub

Private Sub Toggle120Eol_Click()
    Call SetToggles(Me.ActiveControl.Name)
    'Toggle120Exch.caption = "Show ExchIT < 180 days"
    SQLx = ""
    OrderByx = " Order By DaysToEOL "
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "

    If Toggle120Eol.caption Like "*180 Days*" Then
       SQLx = SQLx & " AND DaysToEOL < 181"
       Toggle120Eol.caption = "All Lease Exp"
    Else
        Toggle120Eol.caption = "Lease Exp < 180 Days"
    End If

    If ClientList & "test" <> "test" Then
       SQLx = SQLx & " AND ClientId =" & Me.ClientList & " "
       OrderByx = OrderByx & " ,Client"
    End If

    If SchGrpList & "test" <> "test" Then
       SQLx = SQLx & " AND GroupID =" & Me.SchGrpList & " "
       OrderByx = OrderByx & " ,SchID"
    End If
    SQLz = SQLBase & SQLx & OrderByx
    Me.Form.RecordSource = SQLz
End Sub

Private Sub Toggle120Exch_Click()
    Call SetToggles(Me.ActiveControl.Name)
    Toggle120Eol.caption = "Show EXP < 180 days"
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    If Toggle120Exch.caption Like "*180 Days*" Then

        SQLz = SQLx & " AND DaysToExch Between 1 and 180 Order By Client, SchID, AssetID "
        Toggle120Exch.caption = "Show All ExchangeIT"
    Else
        SQLz = SQLx
        Me.Form.RecordSource = SQLz
        Toggle120Exch.caption = "ExchangeIT < 180 Days"
    End If
    Me.Form.RecordSource = SQLz
    Me.Form.Repaint
End Sub

Private Sub ToggleFindExtd_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    If ToggleFindExtd = True Then SQLz = SQLx & " AND Selected Like '*Extd*' Order By DaysToExch, Client" Else SQLz = SQLx & " Order By DaysToExch, Client"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub ToggleFindNone_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    If ToggleFindNone = True Then SQLz = SQLx & SQLy & " AND Selected = '----  ----  ----' Order By DaysToExch, Client" Else SQLz = SQLx & "Order By DaysToExch, Client"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub ToggleFindTerm_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    If ToggleFindTerm = True Then SQLz = SQLx & SQLy & " AND Selected Like '*Term*' Order By DaysToExch, Client" Else SQLz = SQLx & " Order By DaysToExch, Client"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub ToggleFindExchanged_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    SQLz = SQLx & SQLy & " AND NewSchedule <> '' Order By Client, SchID, NewSchedule"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub ToggleFindNOTExchanged_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    SQLz = SQLx & SQLy & " AND isnull(NewSchedule) Order By NewSchedule, Client"
    Me.Form.RecordSource = SQLz
End Sub

Public Sub EndTypeToShow_Click()
    If Me.EndTypeToShow = 1 Then
        EndType = "EOL"
    Else
        EndType = "EXIT"
    End If
    If EndType = "" Then Call EndTypeToShow_Click
    SQLz = "SELECT Distinct * From " & Queryz & " " & _
           "WHERE AssetID > 0 And " & EndType & "Year > 0 "
    SQLv = SQLv & "ORDER BY " & EndType & "Year," & EndType & "Mth, Client;"
    Me.Form.RecordSource = SQLz
    Me.Detail.Visible = True
    Me.UnitCount.Visible = True
    Me.EOLYear = ""
    Call SetToggles("")
End Sub

Private Sub ExportToExcel_Click()
    Call fExportToExcel(False)
End Sub
Private Sub ExportSendEmail_Click()
    Call fExportToExcel(True)
End Sub
Private Sub SUFEmail_Click()
Call fSufReport
End Sub


Public Function fExportToExcel(Optional MakeEmail As Boolean)

    Dim SQLv As String
    Dim SQLend, Sqlr As String
    Dim XLS, QTRx, TitleX As String
    Dim NotedUnits As Boolean

    SQLend = SplitString(Me.Form.RecordSource, "Where", 1)
    SQLv = ""
    SQLv = SQLv & "Select Client"
    SQLv = SQLv & ",Schedule"
    SQLv = SQLv & ",UnitGroup"
    SQLv = SQLv & ",Assignee"
    SQLv = SQLv & ",LeaseType"
    SQLv = SQLv & ",BDEX"
    SQLv = SQLv & ",BLCD"
    SQLv = SQLv & ",AssetID"
    SQLv = SQLv & ",UnitNum"
    SQLv = SQLv & ",VIN"
    SQLv = SQLv & ",Year"
    SQLv = SQLv & ",Make"
    SQLv = SQLv & ",Model"
    SQLv = SQLv & ",SubType"
    SQLv = SQLv & ",Type"
    SQLv = SQLv & ",CurrRent"
    SQLv = SQLv & ",CurrRentType"
    SQLv = SQLv & ",uStatus"
    SQLv = SQLv & ",[Primary Exp]"
    SQLv = SQLv & ",DaysToEOL"
    SQLv = SQLv & ",[Exit Options]"
    SQLv = SQLv & ",ExtdStart"
    SQLv = SQLv & ",ExtdEnd"
    SQLv = SQLv & ",SufSent"
    SQLv = SQLv & ",SufBack"
    SQLv = SQLv & ",iPassedDate"
    SQLv = SQLv & ",ActualOffLease"
    SQLv = SQLv & ",LeaseRV"
    SQLv = SQLv & ",BookNow As [Book (" & Date & ")]"
    SQLv = SQLv & ",[Book 90d]"
    SQLv = SQLv & ",ODORead"
    SQLv = SQLv & ",ODOSource"
    SQLv = SQLv & ",ODOReadDate"
    SQLv = SQLv & ",Garage"
    SQLv = SQLv & ",ExchDate"
    SQLv = SQLv & ",DaysToExch"
    SQLv = SQLv & ",NewClient"
    SQLv = SQLv & ",NewSchedule"
    SQLv = SQLv & ",NewGroup"
    SQLv = SQLv & ",NewAssetID"
    SQLv = SQLv & ",nInServ"
    SQLv = SQLv & ",ReplacedSchNo"
    SQLv = SQLv & ",ExchangeITMonth"
    SQLv = SQLv & ",ExchangeITType"
    SQLv = SQLv & ",OldSchGrp"
    SQLv = SQLv & ",ClientID"
    SQLv = SQLv & ",SchID"
    SQLv = SQLv & ",GroupID "
    SQLv = SQLv & " FROM ExitOptionsGroupExport "
    SQLv = SQLv & " WHERE " & Trim(SQLend)

    If ToggleQTR1 = True Then QTRx = "-1st Qtr"
    If ToggleQTR2 = True Then QTRx = "-2nd Qtr"
    If ToggleQTR3 = True Then QTRx = "-3rd Qtr"
    If ToggleQTR4 = True Then QTRx = "-4th Qtr"
    QTRx = Me.EOLYear & QTRx
    If EndType = "EOL" Then QTRx = "Units Scheduled to Expire, Report Date " & QTRx Else QTRx = "Units with a 1st ExchangeIT Date " & QTRx
    If ClientList & "test" = "test" Then TitleX = QTRx Else TitleX = "Exit Options For " & Me.Client
    TitleX = TitleX & " - " & Date
    XLS = "|A,BC|C,TC|D,TC|E,BC|F,TC|H,TX|I,BC|K,TC|P,$C|Q,TC|R,BC|S,TC|T,Sint|T,BC|U,TC|Z,TC|"
    XLS = XLS & "AB,$R|AC,$R|AD,$R|AE,SR|AF,TC|AG,TC|AH,TL|AI,TC|AJ,Sint|AJ,TC|"
    XLS = XLS & "AN,TC|AO,TC|AP,TC|AQ,TC|"

    Call MessageUser("Open", "Getting Export", "Just a moment or two...")
    If MakeEmail = True Then
        Dim i As Integer
        Dim msg1  As String
        Dim m2, m3 As String
        Dim Filex As String
        Dim Linkx As String
        Dim subjectx As String
        Dim CcEmailX As String
        Dim ContactEmailX As String
        Dim RST As DAO.Recordset
        msg1 = ""
        m2 = ""
        'Look for Missing OffLease
        Sqlr = SplitString(SQLv, "Order", 0)
        Sqlr = Sqlr & " AND uStatus <>'P' AND uStatus <>'TR' AND uStatus <>'A' AND ActualOffLease IS NULL"
        Sqlr = Sqlr & " ORDER BY Client, AssetID; "

        Set RST = CurrentDb.OpenRecordset(Sqlr, dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        If RST.RecordCount > 0 Then
            NotedUnits = True
            m2 = "<H3>" & RST.RecordCount & " Units Missing the Off Lease Date or have the WRONG Unit Status</H3>"
            m2 = m2 & "<table>"
            m2 = m2 & "<TR style=""background-color:#5CB3FF;"" ALIGN=""Center"">"
            m2 = m2 & "<TH>Client</TH>"
            m2 = m2 & "<TH>Sch|Grp</TH>"
            m2 = m2 & "<TH>AssetID</TH>"
            m2 = m2 & "<TH>Unit Desc</TH>"
            m2 = m2 & "<TH>Rent Type</TH>"
            m2 = m2 & "<TH>uStatus</TH>"
            m2 = m2 & "<TH>Primary</br> Lease Expire</TH>"
            m2 = m2 & "<TH>Days To</br>EOL</TH>"
            m2 = m2 & "<TH>Extd</br>Start</TH>"
            m2 = m2 & "<TH>Extd</br>End</TH>"
            m2 = m2 & "<TH>SUF Sent</TH>"
            m2 = m2 & "<TH>SUF Back</TH>"
            m2 = m2 & "<TH>Insp</br>Passed Date</TH>"
            m2 = m2 & "<TH>Actual</br>Off Lease</TH>"
            m2 = m2 & "</TR>"
            For i = 1 To RST.RecordCount
                If i Mod 2 = 0 Then m2 = m2 & "<TR bgcolor=""WhiteSmoke"">" Else m2 = m2 & "<TR bgcolor=""White"">"
                m2 = m2 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
                m2 = m2 & "<TD ALIGN=""Left"">" & RST!Schedule & "|" & RST!UnitGroup & "</TD>"
                m2 = m2 & "<TD ALIGN=""Left""><b>" & RST!AssetID & "</b></TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!Year & " " & RST!MAKE & " " & RST!MODEL & " " & RST!SubType & " " & RST!Type & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!CurrRentType & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center""><b>" & RST!uStatus & "</b></TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST![Primary Exp] & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center""><b>" & RST!DaysToEOL & "</b></TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!ExtdStart & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!ExtdEnd & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!SUFSent & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!SUFBack & "</TD>"
                m2 = m2 & "<TD ALIGN=""Center"">" & RST!iPassedDate & "</TD>"
                m2 = m2 & "<TD style=""background-color:yellow;"" ALIGN=""Center"">" & RST!ActualOffLease & "</TD>"
                m2 = m2 & "</TR>"
                RST.MoveNext
            Next i
            m2 = m2 & "</table><br /><br />"
            RST.close
            Set RST = Nothing
        End If
        m3 = ""
        Sqlr = ""
        'Look for ALREADY OffLease but Active
        Sqlr = SplitString(SQLv, "Order", 0)
        Sqlr = Sqlr & " AND (uStatus ='P' OR uStatus ='TR' OR uStatus ='A') AND DaysToEOL <=0 AND ActualOffLease IS NULL AND ExtdEnd Is Null"
        Sqlr = Sqlr & " ORDER BY Client, AssetID; "

        Set RST = CurrentDb.OpenRecordset(Sqlr, dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        If RST.RecordCount > 0 Then
            NotedUnits = True
            m3 = "<H3>" & RST.RecordCount & " Units with an Expired Lease but NO Lease Extension Entered</H3>"
            m3 = m3 & "<table>"
            m3 = m3 & "<TR style=""background-color:#5CB3FF;"" ALIGN=""Center"">"
            m3 = m3 & "<TH>Client</TH>"
            m3 = m3 & "<TH>Sch|Grp</TH>"
            m3 = m3 & "<TH>AssetID</TH>"
            m3 = m3 & "<TH>Unit Desc</TH>"
            m3 = m3 & "<TH>Rent Type</TH>"
            m3 = m3 & "<TH>uStatus</TH>"
            m3 = m3 & "<TH>Primary</br> Lease Expire</TH>"
            m3 = m3 & "<TH>Days To</br>EOL</TH>"
            m3 = m3 & "<TH>Extd</br>Start</TH>"
            m3 = m3 & "<TH>Extd</br>End</TH>"
            m3 = m3 & "<TH>SUF Sent</TH>"
            m3 = m3 & "<TH>SUF Back</TH>"
            m3 = m3 & "<TH>Insp</br>Passed Date</TH>"
            m3 = m3 & "<TH>Actual</br>Off Lease</TH>"
            m3 = m3 & "</TR>"
            For i = 1 To RST.RecordCount
                If i Mod 2 = 0 Then m3 = m3 & "<TR bgcolor=""WhiteSmoke"">" Else m3 = m3 & "<TR bgcolor=""White"">"
                m3 = m3 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
                m3 = m3 & "<TD ALIGN=""Left"">" & RST!Schedule & "|" & RST!UnitGroup & "</TD>"
                m3 = m3 & "<TD ALIGN=""Left""><b>" & RST!AssetID & "</b></TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!Year & " " & RST!MAKE & " " & RST!MODEL & " " & RST!SubType & " " & RST!Type & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!CurrRentType & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center""><b>" & RST!uStatus & "</b></TD>"
                m3 = m3 & "<TD style=""background-color:yellow;"" ALIGN=""Center"">" & RST![Primary Exp] & "</TD>"
                m3 = m3 & "<TD style=""background-color:yellow;"" ALIGN=""Center""><b>" & RST!DaysToEOL & "</b></TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!ExtdStart & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!ExtdEnd & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!SUFSent & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!SUFBack & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!iPassedDate & "</TD>"
                m3 = m3 & "<TD ALIGN=""Center"">" & RST!ActualOffLease & "</TD>"
                m3 = m3 & "</TR>"
                RST.MoveNext
            Next i
            m3 = m3 & "</table><br /><br />"
            RST.close
            Set RST = Nothing
        End If
        Filex = "\\fa-poseidon\Atlaas\Temp\ExpiringUnits_180days_" & DatePart("M", Date) & "_" & DatePart("D", Date) & "_" & DatePart("YYYY", Date) & ".xlsx"
        Linkx = ""
        Linkx = Filex
        Call KillFiles("\\fa-poseidon\Atlaas\Temp", "ExpiringUnits*.xlsx")
        subjectx = "Notice: FA Units Expiring Within 180 Days"
        ContactEmailX = "Remarketing@FleetAdvantage.com;TMT@FleetAdvantage.com;hprikes@FleetAdvantage.com;sroyer@FleetAdvantage.com;ncontreras@FleetAdvantage.com"
        'ContactEmailX = "ljohnson@FleetAdvantage.com"
        CcEmailX = "it@FleetAdvantage.com;mdeaguiar@fleetadvantage.com;bantonellis@fleetadvantage.com"

        'make email message
        msg1 = msg1 & "<html><body><H3 style=""color:Navy;font-family:Segoe UI;""><b>Expiring Units =< 180 Days - " & Date & "</H3></b></p>"
        msg1 = msg1 & "<p style=""font-family:Calibri;font-size:11pt"">"
        msg1 = msg1 & "Re: " & subjectx & "<br/><br/>"
        msg1 = msg1 & "All,<br/><br/>"
        msg1 = msg1 & "Please review the attached spreadsheet: (<B>" & SplitString(Filex, "Temp\", 1) & "</B>)<br />"
        msg1 = msg1 & "It shows a list of Equipment that are set to expire, within 180 days.<br /><br />"
        msg1 = msg1 & "All or any required actions are based on your Team`s processes for expiring units.<br />"
        If NotedUnits = True Then msg1 = msg1 & "<H3>Units that need attention:</H3>"
        msg1 = msg1 & m2
        msg1 = msg1 & m3
        msg1 = msg1 & "If you see any errors or any<i> out of place items</i> within the list of attached units, please contact the IT Department.<br /><br />"
        msg1 = msg1 & "Regards,<br />"
        msg1 = msg1 & "</body></html>"
        SendReport2Excel SQLv, "EXIT Options", TitleX, True, XLS, 0, False, 1, False, , Filex
        ''Call SendEmailMessage(ContactEmailX, subjectx, msg1, False, True, False, CcEmailX, "", "", Linkx)
        Call SendServerEmail(ContactEmailX, CcEmailX, subjectx, msg1, "ljohnson@fleetadvantage.com", Linkx)
    Else
        SendReport2Excel SQLv, "EXIT Options", TitleX, True, XLS, 0, False
    End If
    Call MessageUser("Close")
End Function

Private Sub EOLYear_Click()
    If EndType = "" Then Call EndTypeToShow_Click
    SQLz = "SELECT Distinct * From " & Queryz & " " & _
           "WHERE AssetID > 0 And " & EndType & "Year = " & Me.EOLYear & " " & _
           "ORDER BY Client, ReplacedSchNo;"
    Me.Form.RecordSource = SQLz
    Me.Detail.Visible = True
    Me.UnitCount.Visible = True
    Call SetToggles("")
End Sub

Private Sub ToggleQTR1_Click()
Call SetQtr(1)
End Sub

Private Sub ToggleQTR2_Click()
Call SetQtr(2)
End Sub

Private Sub ToggleQTR3_Click()
Call SetQtr(3)
End Sub

Private Sub ToggleQTR4_Click()
Call SetQtr(4)
End Sub
Private Sub SetQtr(QuarterX As Integer)
    If Me.EOLYear & "test" = "test" Then
        Me("ToggleQTR" & QuarterX) = False
        Exit Sub
    End If
    Call SetToggles("ToggleQTR" & QuarterX)
    If Me("ToggleQTR" & QuarterX) = True Then
    SQLz = "SELECT Distinct * From " & Queryz & " " & _
           "WHERE AssetID > 0 And " & EndType & "Year = " & Me.EOLYear & " And " & EndType & "QTR = " & QuarterX & " " & _
           "ORDER BY Client, ReplacedSchNo;"
    Else
    SQLz = "SELECT Distinct * From " & Queryz & " " & _
           "WHERE AssetID > 0 And " & EndType & "Year = " & Me.EOLYear & " " & _
           "ORDER BY Client, ReplacedSchNo;"
    End If
    Me.Form.RecordSource = SQLz
    Me.Detail.Visible = True
    Me.UnitCount.Visible = True
End Sub

Private Sub ToggleFindOLAP_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    If ToggleFindOLAP = True Then SQLz = SQLx & SQLy & " AND Selected Like '*OVLP*' Order By DaysToExch, Client" Else SQLz = SQLx & "Order By DaysToExch, Client"

    Me.Form.RecordSource = SQLz
End Sub

Private Sub ToggleOffLease_Click()
    Call SetToggles(Me.ActiveControl.Name)
    SQLx = ""
    SQLBase = "Select Distinct * From " & Queryz & " Where AssetID > 0 "
    If SchGrpList & "test" <> "test" Then
        SQLx = SQLBase & " AND GroupID =" & Me.SchGrpList & " "
        GoTo MakeSQL
    Else
        SQLx = SQLBase
    End If
    If ClientList & "test" <> "test" Then SQLx = SQLBase & " AND ClientId =" & Me.ClientList & " " Else SQLx = SQLBase
MakeSQL:
    SQLy = ""
    If Me.Toggle120Eol <> 0 Then SQLy = " AND DaysToEOL < 181"
    If Me.Toggle120Exch = True Then SQLy = " AND DaysToExch Between 1 and 180"
    If ToggleOffLease = True Then SQLz = SQLx & SQLy & " AND isdate(ActualOffLease) Order By ActualOffLease, Client" Else SQLz = SQLx & "Order By DaysToExch, Client"

    Me.Form.RecordSource = SQLz
End Sub
Public Function CheckStatus() As Boolean
Dim teststatus As Boolean
    If Me.uStatus = "A" Then teststatus = 1
    If Me.uStatus = "TR" Then teststatus = 1
    If Me.uStatus Like "E*" Then teststatus = 1
    If teststatus = False Then
    MsgBox "An EXIT Option cannot be applied to a '" & Me.uStatus & "' Status unit." & _
               vbNewLine & "      If you need to apply an EXIT Option," & _
               vbNewLine & "      Please update the status of the Unit." & _
               vbNewLine & "         OR" & _
               vbNewLine & "      Create an IT Department Ticket", , "EXIT Option Create Warning"
    End If
    CheckStatus = teststatus
End Function
Private Sub ModifyExch_Click()
       If Not FAUserAccessGroup Like "*TMT*" Or Not FAUserAccessLevel > 50 Then Exit Sub
       If CheckStatus = False Then Exit Sub
    Call MakeExitOptionUnit(Me.UnitID, Me.SchID, Me.groupID, "EXchangeIT")
End Sub

Private Sub ModifyEXTD_Click()
    If Not FAUserAccessGroup Like "*TMT*" Or Not FAUserAccessLevel > 50 Then Exit Sub
    If CheckStatus = False Then Exit Sub
    If IsNull(Me.UnitID) Then Exit Sub
    Call MakeExitOptionUnit(Me.UnitID, Me.SchID, Me.groupID, "Extend")
End Sub

Private Sub ModifyTERM_Click()
    If Not FAUserAccessGroup Like "*TMT*" Or Not FAUserAccessLevel > 50 Then Exit Sub
    If CheckStatus = False Then
        MsgBox "An EXIT Option may only applied to an A or TR unit." & vbNewLine & "  If you need to apply an EXIT Option," & vbNewLine & "  Please update the status of the Unit." & vbNewLine & "    or" & vbNewLine & "  Create a Ticket with IT", , "EXIT Option Create Warning"
        Exit Sub
    End If
    If IsNull(Me.UnitID) Then Exit Sub
    Call MakeExitOptionUnit(Me.UnitID, Me.SchID, Me.groupID, "Terminate")
End Sub

Private Sub ModifyHIST_Click()
    If Not FAUserAccessGroup Like "*TMT*" Or Not FAUserAccessLevel > 50 Then Exit Sub
    Dim SchIDz As Integer
    SchIDz = Me.SchID
    If DCount("UnitID", "UnitExitOption", "UnitID=" & Me.UnitID) < 1 Then
        MsgBox "AssetID: " & Me.UnitID & " does not have an EXIT Option applied nor any EXIT Option history." & vbNewLine & "To save time, the Exit Options History Form will NOT be loaded." & vbNewLine & vbNewLine & "If the Unit should have EXIT Option history, please create a Ticket with IT.", , "Unit has No History"
        Exit Sub
    End If
    Call MakeExitOptionUnit(Me.UnitID, Me.SchID, Me.groupID, "History")
End Sub

Private Sub NewUnit_DblClick(Cancel As Integer)
    If Me.NewUnit & "test" = "test" Then Exit Sub
    Call OpenUnitsForm(Me.NewUnit)
    Me.UnitUnitNum.SetFocus
End Sub

Private Sub SearchClient_Change()
    SQLy = "SELECT ClientID, Client, [Type] " & _
           "FROM ClientMasterLeasesTypeCount_Summed Where Client Like '*" & Me.SearchClient.Text & "*' " & _
           "ORDER BY Client;"
    Me.ClientList.RowSource = SQLy
    Me.SearchML = ""
End Sub

Private Sub SearchML_Change()
    Me.SearchClient = ""
        WhereX = " WHERE MLOrig ='FA' "
        If Not IsNull(Me.ClientList.Column(0)) Then WhereX = WhereX & "AND ClientID=" & Me.ClientList.Column(0) & " "
        WhereX = WhereX & "AND (Schedule Like '*" & Me.SearchML.Text & "*' OR UnitGroup Like '*" & Me.SearchML.Text & "*') "
        Me.SchGrpList.RowSource = SQLGrp & WhereX & GroupBY & ";"
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Forms!Unitsform.UnitsTabControl.Value = 8
    Me.UnitNum.SetFocus
End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
    Call UnitID_DblClick(1)
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "MEdit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
    Me.SchGrpList = Null
    Me.SchGrpList.Enabled = False
    'Me.UnitCount.Visible = False
End Sub

Public Sub ClearList(ListBoxName As String)
    Dim ItemsX As Variant
    With Me(ListBoxName)
        For Each ItemsX In .ItemsSelected
            .Selected(ItemsX) = False
        Next ItemsX
        Me(ListBoxName) = Null
    End With
End Sub

Public Sub SetFilters()
    Toggle120Eol.Value = False
    Toggle120Exch.Value = False
    Toggle120Eol.caption = "Lease Exp < 180 Days"
    Toggle120Exch.caption = "ExchangeIT < 180 Days"
End Sub
