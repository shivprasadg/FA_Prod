Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Dim SQLz   As String
    Dim XLS   As String
Private Sub CompanyName_Click()
    MessageUser "Open", "Opening Client's Form", "just a moment.."
    DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
    MessageUser "Close"
End Sub



Private Sub ClientList_Click()

    SQLz = "SELECT * FROM vw_UnitGroupHealth "
    SQLz = SQLz & "WHERE ClientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & "AND [Assigned] <>'F' AND [Funded]='N'"
    SQLz = SQLz & "OR (ClientID=" & Me.ClientList.Column(0) & " AND [Funded]='Y' AND [Assigned]='P' ) "
    SQLz = SQLz & "ORDER BY SchID Desc;"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub CustomerDeliveryReportNew_Click()
   Call CustDelvReportFinal
End Sub

Private Sub ExcelLegal_Click()
   Call SchHealthReport("Legal")
End Sub

Private Sub Form_Load()

End Sub

Private Sub OpenMassMod_Click()
    If Me.groupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    MessageUser "Open", "Opening Mass-Mod Form", "just a moment.."
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = Me.ClientList
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = DLookup("MLID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = DLookup("SchID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = Me.groupID
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
     MessageUser "Close"
End Sub
Private Sub Purchasing_Click()
    Call StartSignOff(Me.groupID, "SignOff1")
    Me.Form.Requery
End Sub
Private Sub InvIn_Click()
    Call StartSignOff(Me.groupID, "SignOff2")
    Me.Form.Requery
End Sub

Private Sub MSOValid_Click()
    Call StartSignOff(Me.groupID, "SignOff3")
    Me.Form.Requery
End Sub
Private Sub Operations_Click()
    Call StartSignOff(Me.groupID, "SignOff4")
    Me.Form.Requery
End Sub

Private Sub ShowAll_Click()
    SQLz = "SELECT * FROM vw_UnitGroupHealth "
    SQLz = SQLz & "Where [Assigned] <>'F' AND [Funded]='N'"
    SQLz = SQLz & "ORDER BY ClientShortName, Schedule Desc;"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub SignOffsPendingBtn_Click()
    Call SignOffsPendingReport
End Sub

Private Sub TMT_Click()
    Call StartSignOff(Me.groupID, "SignOff5")
    Me.Form.Requery
End Sub

Private Sub Docs_Click()
    Call StartSignOff(Me.groupID, "SignOff6")
    Me.Form.Requery
End Sub
Private Sub Syndication_Click()
    Call StartSignOff(Me.groupID, "SignOff7")
    Me.Form.Requery
End Sub
Private Sub Client_DblClick(Cancel As Integer)
    If Me.groupID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & DLookup("SchID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
End Sub
Private Sub SchGrp_DblClick(Cancel As Integer)
    If Me.groupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.groupID
End Sub

Private Sub UnFundedOECSch_Click()

    SQLz = "Select ClientShortName as Client, Schedule"
    SQLz = SQLz & ", GroupsNoinSchedule(S.SchID) As GroupsIn"
    SQLz = SQLz & ",Count(AssetID) as Units"
    SQLz = SQLz & ",Min(S.UnitDescShortVer) as uType"
    SQLz = SQLz & ",IIf(Max(S.UnitDescShortVer) = Min(S.UnitDescShortVer),'', Max(S.UnitDescShortVer)) as uType2"
    SQLz = SQLz & ",Cint(avg(S.LeaseTermPrimary)) As AvgTerm"
    SQLz = SQLz & ",Min(NZ(ActualDeliveryDate,EsitmatedDeliveryDate)) as FirstDelvDate"
    SQLz = SQLz & ",Max(NZ(ActualDeliveryDate,EsitmatedDeliveryDate)) as LastDelvDate"
    SQLz = SQLz & ",BankShortName As Assignee"
    SQLz = SQLz & ",Max(G.LeaseType) as LeaseType"
    SQLz = SQLz & ",Sum(OECUnit) as OEC"
    SQLz = SQLz & ",Max(NZ(iBank,'')) as [Interim Bank]"
    SQLz = SQLz & ",Sum(NZ(UnitIFAmt,0)) as [Interim Amt]"  ' N
    SQLz = SQLz & ",IIF(MIN(UnitIFBankID) = 4321 AND MAX(UnitIFBankID) = 4321,'', DateAdd(""m"",6,Min([UnitIFDate]))) as [Interim Expire]"
    SQLz = SQLz & ",Min(U.UnitPayOff) as [Interim Paid]"      ' P
    SQLz = SQLz & ",Min(Assigned_Indicator) As [Fund Status]"
    SQLz = SQLz & " From ((vw_SixKeys As S"
    SQLz = SQLz & " Inner Join Units as U on S.AssetId = U.UnitID)"
    SQLz = SQLz & " Inner Join SchGrp as G on S.SgrpId =G.sGrpId)"
    SQLz = SQLz & " LEFT Join (Select BankShortName As iBank, BankID FROM Banks) as B on U.UnitIFBankID = B.BankID"
    SQLz = SQLz & " Where MLorig='FA' and FundedDate is Null"
    SQLz = SQLz & " Group By ClientShortName, Schedule, BankShortName ,iBank, SCHID"
    SQLz = SQLz & " Order By ClientShortName, Schedule"

    XLS = "|A,DD3|C,W20|D,BC|F,W15|G,BC|L,$E|M,W09|N,$E|N,TC|O,W10|O,TR|P,W10|Q,BC|Q,W06|"
    XLS = XLS & "A,Wrap|M,cColor:3:191.143.0|N,cColor:3:191.143.0|O,cColor:3:191.143.0|P,cColor:3:191.143.0|"
    Call MessageUser("Open", "Getting Schedule OEC Report", "Drop it like it's hot...")
    SendReport2Excel SQLz, "UnFunded", "UnFunded Schedules - " & Date, True, XLS, 0, False
    Call MessageUser("Close")
End Sub
