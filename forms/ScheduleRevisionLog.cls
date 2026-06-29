Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, SQLx, GroupIDx, SchIDx  As String
Dim Ccrl       As Control

Private Sub Action_Date_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_ActionDate = SplitString(Me.Action_Date, " ", 0)
    Call SetRecordSource("AND ActionDate Like '*" & Me.Search_ActionDate & "*' ")
End Sub

Private Sub Action_Taken_Click()
Call fieldname_Click
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
On Error GoTo FixMe
If Me.FieldGroup = "Reasons" Then
Dim PostionX As String
PostionX = Me.identity
Me.Parent.RevList = Me.identity
Me.Parent.RevList_Click
'Me.Form.Recordset.Move PostionX
Set RST = Me.RecordsetClone
       RST.FindFirst "identity = " & PostionX & ""
       Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End If
FixMe:
End Sub

Private Sub Action_Date_Click()
Call fieldname_Click
End Sub

Public Sub Form_Current()
    SchIDx = Me.SchID
    GroupIDx = Me.groupID
    Me.txtCurrent = Nz(Me.ID, 0)
End Sub
Public Function SetRecordSource(Optional WhereX As String)
On Error GoTo FixMe
    SQLz = ""
    SQLz = SQLz & "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev, Nz([NewValue],0) AS New, " & _
           "SysUser AS [User], FieldName, GroupID, SchID, RevID, ID " & _
           "FROM vw_FAID_RevScheduleUnion "
    SQLz = SQLz & "WHERE SchID=" & SchIDx & " " & WhereX & " ORDER BY ActionDate DESC;"
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

Private Sub New_Click()
Call fieldname_Click
End Sub

Private Sub Prev_Click()
Call fieldname_Click
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

Private Sub SendEmail_Click()
Call SendATFEmail(Me.SchID, Nz(Me.groupID, 0), Me.txtCurrent)
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

Private Sub User_DblClick(Cancel As Integer)
    Call ClearSearchBoxes("All")
    Me.Search_SysUser = Me.User
    Call SetRecordSource("AND SysUser =    '" & Me.Search_SysUser & "' ")
End Sub

Private Sub ViewActiveATFReport_Click()
On Error GoTo FixMe
If Me.SchID & "test" = "test" Then Exit Sub
    Dim FAIDx As String
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
    Dim RST As Object
Set RST = Me.RecordsetClone
On Error Resume Next
RST.MoveLast
On Error GoTo 0
Me.Events.Value = RST.RecordCount
RST.close
Set RST = Nothing
End Sub
