Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, SQLx, GroupIDx As String
Dim Ccrl       As Control
Private Sub Prev_Click()
    Call fieldname_Click
End Sub
Private Sub New_Click()
    Call fieldname_Click
End Sub
Private Sub ActionDate_Click()
    Call fieldname_Click
End Sub
Private Sub ActionTaken_Click()
    Call fieldname_Click
End Sub

Private Sub Send_EmailAttachement_Click()
'Call SendATFEmail(Me.SchID, Me.groupID, Me.SelectedID)
End Sub
Public Sub SendATFEmail(SchIDx As Long, GroupIDx As Long, EventID As Long)
    Dim subjectx, ClientX, SQLz, SQLx As String
    Dim msg1, NoticeTypeX, Schx, Grpx, FontSetting, UnitDesX As String
    Dim x As Integer
    Dim RST    As DAO.Recordset
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SchID=" & SchIDx)
    Schx = DLookup("Schedule", "vw_SixKeys", "SchID=" & SchIDx)
    If GroupIDx = 0 Then Grpx = DLookup("UnitGroup", "vw_SixKeys", "SchID=" & SchIDx) Else Grpx = DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & GroupIDx)
    NoticeTypeX = DLookup("FieldGroup", "ATF_ChangeReasons", "INDX=" & EventID)
    If IsNull(DLookup("UnitID", "ATF_ChangeReasons", "INDX=" & EventID)) Then
    UnitDesX = "N/A"
    Else
    UnitDesX = Nz(DLookup("UnitDescShortVer", "vw_SixKeys", "UnitID=" & DLookup("UnitID", "ATF_ChangeReasons", "INDX=" & EventID)), "N/A")
    End If
    SQLz = "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev,Nz([NewValue],0) AS New, SysUser AS [User], FieldName, " & _
           "GroupID, SchID, FieldGroup, RevID, ID FROM vw_FAID_RevScheduleUnion WHERE ID = '" & EventID & "';"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount = 0 Then Exit Sub
    subjectx = "Schedule Revision Number:" & RST!RevID & "; " & ClientX & " Sch " & Schx & "|Grp " & Grpx
    'make email message
    FontSetting = "<span style='font-size:10pt;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>"
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'><b> Changes to " & RST!FieldName & "</b></p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:13px'>To: ATF Distribution<br /><br />"
    msg1 = msg1 & "Re. Client:  <B>" & ClientX & "</B><br />"
    msg1 = msg1 & HSpaces(6) & "MLNo: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><Br />"
    msg1 = msg1 & HSpaces(6) & "Schedule: <B>" & Schx & "</B>"
    msg1 = msg1 & HSpaces(4) & "Group: <B>" & Grpx & "</B> "
    msg1 = msg1 & HSpaces(4) & "Total Units in Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><Br />"
    msg1 = msg1 & HSpaces(6) & "Date changes made: <B> " & RST!ActionDate & "</B> "
    msg1 = msg1 & HSpaces(6) & "Changes made by: <B> " & RST!User & "</B><Br />"
    msg1 = msg1 & "<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=2>"
    msg1 = msg1 & "<TR ALIGN='Center'>"
    msg1 = msg1 & FontSetting
    msg1 = msg1 & "<TH>Event</TH><TH>From</TH><TH>To</TH><TH>Current Unit Desc</TH>"
    msg1 = msg1 & "<TH>Reason Provided</TH>"
    msg1 = msg1 & "</TR></span>"
    RST.MoveFirst
    For x = 1 To RST.RecordCount
   '     If NoticeIDx < 12 Then OldX = Format(Rst!OldValue, "$#,#.#0") Else OldX = Rst!OldValue
   '     If NoticeIDx < 12 Then NewX = Format(Rst!NewValue, "$#,#.#0") Else NewX = Rst!NewValue
        msg1 = msg1 & FontSetting
        msg1 = msg1 & "<TD ALIGN='Left'>" & RST!FieldName & "</TD>"
        msg1 = msg1 & "<TD ALIGN='Right'>" & RST!Prev & "</TD>"
        msg1 = msg1 & "<TD ALIGN='Right'>" & RST!New & "</TD></span>"
        msg1 = msg1 & FontSetting
        msg1 = msg1 & "<TD ALIGN='Center'>" & UnitDesX & "</TD>"
        msg1 = msg1 & "<TD ALIGN='Left'>" & RST!ActionDesc & "</TD>"
        msg1 = msg1 & "</span></TR>"
        RST.MoveNext
    Next x
    msg1 = msg1 & "</table>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Dim EmailToList As String
    EmailToList = ""
    'EmailToList = "ATFDistribution@FleetAdvantage.com"
    Call SendEmailMessage(EmailToList, subjectx & "", msg1, False, True, False, "ljohnson@fleetadvantage.com")
    'Mark Email Sent in the notifaction log
    'DoCmd.SetWarnings False
    'DoCmd.RunSQL ("Update ATLAASNotifications Set NoticeSent=1 Where INDX=" & MaxINDXLog & ";")
    'DoCmd.SetWarnings True
End Sub


Private Sub User_Click()
    Call fieldname_Click
End Sub

Private Sub ActionDate_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_ActionDate = SplitString(Me.ActionDate, " ", 0)
    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDate & "*' ")
End Sub

Private Sub AddRevision_Click()
On Error GoTo FixMe
    DoCmd.OpenForm "Schedule_Form", acNormal, , "SchId=" & Me.SchID
    Exit Sub
FixMe:
    MsgBox "There are no events for this Schedule... yet", vbCritical, AtlaasVersion
End Sub

Private Sub ClearResetForm_Click()
    Call ClearLabel_Click
End Sub

Private Sub ClearLabel_Click()
    On Error Resume Next
    Me.Parent.UnitRevDetail.RowSource = ""
    Me.Parent.UnitRevDetail.Visible = False
    On Error GoTo 0
    Call ClearSearchBoxes("All")
    Call SetRecordSource
End Sub

Private Sub FieldName_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_FieldName = Me.FieldName
    Call SetRecordSource("AND FieldName = '" & Me.Search_FieldName & "' ")
End Sub
Private Sub fieldname_Click()
    On Error GoTo 0
    Me.Parent.UnitRevDetail.RowSource = ""
    Me.Parent.UnitRevDetail.Visible = False
    Me.Parent.ReasonAction.Visible = False
    Me.Parent.ReasonNote.Visible = False
    Me.Parent.UnitRevDetailLabel.caption = Me.FieldName & " Detail"
    Me.Parent.ReasonAction = ""
    Me.Parent.ReasonNote = ""
    Me.Parent.ReasonINDX = ""

    If Me.FieldGroup = "UnitSpec" _
       Or Me.FieldGroup = "Rent" _
       Or Me.FieldGroup = "Residual" _
       Or Me.FieldGroup = "Vendor" _
       Or Me.FieldGroup = "Economics" _
       Or Me.FieldGroup = "Structure" _
       Or Me.FieldGroup = "Terms" Then

        Dim FieldGroupX, DateTimeStart, DateTimeEnd, JustDate, ViewX As String
        Dim UnitIDStartLng As Long
        JustDate = SplitString(Me.ActionDate, " ", 0)
        If Len(JustDate) < 10 Then JustDate = "0" & JustDate
        If Len(JustDate) < 10 Then JustDate = Left(JustDate, 3) & "0" & Right(JustDate, 6)
        SQLx = "SELECT AssetID, Ccur(OldValue) As Prev, Ccur(NewValue) As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "

        If Me.FieldName Like "*Rent*" Or Me.FieldName Like "*Cost*" Or Me.FieldName Like "*Resi*" Then
            If Me.FieldName Like "*Rent*" Or Me.FieldName Like "*Cost*" Then
                FieldGroupX = "Rent Change"
                ViewX = "vw_FAID_RevRents"
                UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            End If
        End If

        If Me.FieldName Like "*Resi*" Then
            FieldGroupX = "Residual Change"
            ViewX = "vw_FAID_RevResidual"
            UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
        End If

        If Me.FieldName Like "*Vendor*" Or Me.FieldName Like "*delivery*" Then
            FieldGroupX = "Vendor Change"
            ViewX = "vw_FAID_RevUnitVendor"
            UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If

        If Me.FieldGroup = "UnitSpec" Then
            FieldGroupX = "Unit Spec Change"
            ViewX = "vw_FAID_RevUnitSpec"
            UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If
        If Me.FieldGroup = "Structure" Then
            FieldGroupX = "Structure"
            ViewX = "vw_FAID_RevMoveGrps"
            UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT AssetID, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldName, sysUser, ActionDate, GroupID "
        End If
        DateTimeStart = Left(Me.ActionDate, Len(Me.ActionDate) - 5) & "00 " & Right(Me.ActionDate, 2)
        DateTimeEnd = Left(Me.ActionDate, Len(Me.ActionDate) - 5) & "02 " & Right(Me.ActionDate, 2)
        Me.Parent.UnitRevDetail.Visible = True
        Me.Parent.ReasonAction.Visible = True
        Me.Parent.ReasonNote.Visible = True
        Me.Parent.UnitRevDetail.ColumnWidths = "0.6in;.9in;.9in;0in;0;1in;0"
        'RealFieldX = DLookup("FAIDField", "ATF_FAID_FieldMatrix", "FieldLabel='" & Me.FieldName & "'")

        SQLx = SQLx & "FROM  " & ViewX & " " & _
               "WHERE GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' " & _
               "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By AssetID, ActionDate Desc;"

        If Me.FieldGroup = "Terms" Then
            FieldGroupX = "Terms Change"
            ViewX = "vw_FAID_RevTerms"
            ' UnitIDStartLng = Nz(DMin("AssetID", ViewX, "GroupID=" & Me.Parent.SGrpID & " AND FieldName='" & Me.FieldName & "' AND ActionJustDate = '" & JustDate & "'"), 0)
            SQLx = "SELECT FieldName, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldType, sysUser, ActionDate, GroupID " & _
                   "FROM  vw_FAID_RevTerms " & _
                   "WHERE GroupID=" & Me.Parent.SGrpID & " " & _
                   "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By ActionDate Desc;"
            Me.Parent.UnitRevDetail.ColumnWidths = "1.3in;0.7in;0.7in;0in;0;1in;0"
        End If
        If Me.FieldName Like "*Swaps*" Then
            Dim EcoType As String
            ViewX = "vw_FAID_RevEconomics"
            If Me.FieldName = "Sales Swaps" Then EcoType = "FieldName Like '*(sales)*'"
            If Me.FieldName = "Synd Swaps" Then EcoType = "Not FieldName Like '*(sales)*'"
            SQLx = "SELECT FieldName, OldValue As Prev, NewValue As New, '" & Me.FieldName & "' As FieldType, sysUser, ActionDate, GroupID " & _
                   "FROM  vw_FAID_RevEconomics " & _
                   "WHERE GroupID=" & Me.Parent.SGrpID & " AND " & EcoType & " " & _
                   "AND DateValue(ActionDate) =  DateValue('" & Me.ActionDate & "') Order By ActionDate Desc;"
            Me.Parent.UnitRevDetail.ColumnWidths = "1.3in;0.7in;0.7in;0in;0;1in;0"
        End If

        Me.Parent.UnitRevDetail.RowSource = SQLx
        SQLx = ""
GetChangeReason:
        SQLx = SQLx & "Select * From ATF_ChangeReasons "
        SQLx = SQLx & "WHERE CDate(SplitString(ChangeDate,' ',0))='" & CDate(JustDate) & "' "
        SQLx = SQLx & "AND [FieldName]='" & Me.FieldName & "' "
        'Check if this was a 'unit level field' event
        If Me.FieldGroup <> "Economics" And Me.FieldGroup <> "Terms" Then
            SQLx = SQLx & "AND [UnitID]='" & UnitIDStartLng & "'"
        Else
            SQLx = SQLx & ";"
        End If
        Dim RST As DAO.Recordset
        On Error GoTo FixMe
        Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        On Error Resume Next
        If RST.RecordCount < 1 Then Exit Sub
        RST.MoveLast
        RST.MoveFirst
        If RST.RecordCount < 1 Then Exit Sub
        On Error GoTo 0
        RST.MoveFirst
        Me.Parent.ReasonAction = RST!Action
        Me.Parent.ReasonNote = RST!Reason
        Me.Parent.ReasonINDX = RST!INDX
        RST.close
        Set RST = Nothing
    End If
Exit Sub
FixMe:
Set RST = Nothing
'MsgBox "There was a SQL on this event", vbCritical, AtlaasVersion
End Sub
Public Sub Form_Current()
  GroupIDx = Me.GroupID
  Me.SelectedID = Me.identity
  Me.txtCurrent = Me.FieldName
End Sub
Public Function SetRecordSource(Optional WhereX As String)
    On Error GoTo FixMe
    SQLz = ""
    SQLz = SQLz & "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev, Nz([NewValue],0) AS New, " & _
           "SysUser AS [User], FieldName, GroupID, ATFGroupSpecID, SchID, FieldGroup " & _
           "FROM vw_FAID_RevGroupUnion "
    SQLz = SQLz & "WHERE GroupID=" & GroupIDx & " " & WhereX & " ORDER BY ActionDate DESC;"
    Me.Form.RecordSource = SQLz
    Call GetCount
FixMe:
End Function
Private Sub Form_Load()
    On Error Resume Next
    Me.Parent.UnitRevDetail.RowSource = ""
    Me.Parent.UnitRevDetail.Visible = False
    On Error GoTo 0
    Call ClearSearchBoxes("All")
    Call GetCount
End Sub

Private Sub Search_ActionTaken_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND ActionDesc Like '*" & Me.Search_ActionTaken & "*' ")
End Sub

Private Sub Search_FieldName_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND FieldName = '" & Me.Search_FieldName & "' ")
End Sub

Private Sub Search_New_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND NewValue Like '*" & Me.Search_New & "*' ")
End Sub
Private Sub Search_Prev_AfterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND OldValue Like '*" & Me.Search_Prev & "*' ")
End Sub
Private Sub Search_SysUser_Click()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND SysUser =    '" & Me.Search_SysUser & "' ")
End Sub

Private Sub Search_ActionDate_afterUpdate()
    Call ClearSearchBoxes(Me.ActiveControl.Name)
    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDate & "*' ")
End Sub

Private Sub ClearSearchBoxes(Excludeme As String)
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Search" And Me(Ccrl.Name).Name <> Excludeme Then
            Me(Ccrl.Name) = ""
        End If
    Next Ccrl
End Sub



Private Sub User_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_SysUser = Me.User
    Call SetRecordSource("AND SysUser =    '" & Me.Search_SysUser & "' ")
End Sub

Private Sub ViewActiveATFReport_Click()
    On Error GoTo FixMe
    If Me.SchID & "test" = "test" Then Exit Sub
    Dim FAIDx  As String
    FAIDx = Nz(DLookup("FAID", "vw_SixKeys", "MLOrig= 'FA' and Schid=" & Me.SchID), "")
    If FAIDx = "" Then
        MsgBox "This NOT an FA Leased Schedule", vbCritical, AtlaasVersion
        Exit Sub
    Else
      Call OpenActiveATFReport(Me.SchID)
    End If
FixMe:
End Sub
Private Sub GetCount()
    Dim RST    As Object
    Set RST = Me.RecordsetClone
    On Error Resume Next
    RST.MoveLast
    On Error GoTo 0
    Me.Events.Value = RST.RecordCount
    RST.close
    Set RST = Nothing
End Sub
