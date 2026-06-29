Attribute VB_Name = "ATF_Functions"
Option Compare Database
Option Explicit
Public SessionInt As Long
Public Function CycleForReasons(GroupIDl As Long)
    Call SyncAuditTables
    If Nz(GroupIDl, 0) = 0 Or Nz(GroupIDl, "") = "" Then Exit Function
    Call MessageUser("Open", "Checking For Changes...", "Audit Table Stuff")
    Call CheckATFChanges(GroupIDl, "Rent Change")
    Call CheckATFChanges(GroupIDl, "Residual Change")
    Call CheckATFChanges(GroupIDl, "Unit Spec Change")
    Call CheckATFChanges(GroupIDl, "Unit Vendor Change")
    Call CheckATFChanges(GroupIDl, "ChildPart Change")
    'Disabled 8/13/2020 per OPS - Pat
    'If AllUnitsInGroup_Invoiced(Me.GrpID) = True Then Call AllUnitsInGroup_InvoicedMessage(Me.GrpID)
    'If DMax("Status", "vw_SixKeys", "SGrpID=" & GroupIDl) = "A" Then Call AllUnitsInGroup_AcceptedMessage(Me.GrpID)
    Call MessageUser("Close")
End Function

Public Function CheckATFChanges(GroupIDx As Long, NoticeTypeX As String, Optional ReportDate As String, _
    Optional FormNameX As String, Optional TestEvo As Boolean, Optional DoSync As Boolean)
    '*****************************
    ' To test the module         *
    ' set TestEvo to True        *
    ' TestEvo = True
    '*****************************
    Dim msg1   As String
    Dim SQLz   As String
    Dim ClientX, Schx, AllAccepted, AllPDSet, SQLx, SQLy, Grpx, FontSetting, Msgx, OldX, NewX As String
    Dim ReplyX As String
    Dim RevQuery As String
    Dim PDFfileName As String
    Dim InputText As String
    Dim NoticeIDx As Integer
    Dim ClientIDint, x, UnitSpecX, ATFIDx, Finalized As Integer
    Dim SchIDInt As Long
    Dim UnitIDInt As Long
    Dim EventIDint As Long
    Dim RevIDint As Long
    Dim UserStopped As Boolean
    Dim RST    As DAO.Recordset
    '________________
    'Gets Today's Audit into Processing
    If DoSync <> False Then Call SyncAuditTables
    '________________

    SQLz = "Select CompanyName, clientID, Schedule, SchID, UnitGroup " & _
    "From vw_SixKeys WHERE SGrpID=" & GroupIDx & " " & _
    "Group By CompanyName, clientID, Schedule, SchID, UnitGroup;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    Schx = RST!Schedule
    Grpx = RST!UnitGroup
    SchIDInt = RST!SchID
    ClientX = RST!CompanyName
    ClientIDint = RST!ClientID
    RST.close
    If ReportDate & "test" = "test" Then ReportDate = Date
    If Len(ReportDate) < 10 And DatePart("m", CDate(ReportDate)) < 10 Then ReportDate = "0" & ReportDate
    If Len(ReportDate) < 10 Then ReportDate = Left(ReportDate, 3) & "0" & Right(ReportDate, 6)
    If FormNameX & "test" = "test" Then FormNameX = "Mass-Mod"
    SQLz = ""
    PDFfileName = ""
    'See what type of change requested
    Select Case NoticeTypeX
        Case "Rent Change"
            RevQuery = "vw_FAID_RevRentGrouped"
            NoticeIDx = 10
        Case "Residual Change"
            RevQuery = "vw_FAID_RevResidualGrouped"
            NoticeIDx = 11
        Case "Unit Vendor Change"
            RevQuery = "vw_FAID_RevUnitVendorGrouped"
            NoticeIDx = 12
        Case "Unit Spec Change"
            RevQuery = "vw_FAID_RevUnitSpecGrouped"
            NoticeIDx = 13
        Case "Move Units"
            RevQuery = "vw_FAID_RevMoveGrpsGrouped"
            NoticeIDx = 14
            FormNameX = "Move Units"
        Case "Lease Terms Change"
            RevQuery = "vw_FAID_RevTerms"
            NoticeIDx = 15
            FormNameX = "Group Form"
        Case "Economics Change"
            RevQuery = "vw_FAID_RevEconomics"
            NoticeIDx = 16
            FormNameX = "Group Form"
            SQLx = "Select * From vw_FAID_RevEconomics WHERE GroupID=" & GroupIDx & " AND FieldGroup = 'Economics' AND ActionJustDate = '" & ReportDate & "' AND ATFField=True ;"
        Case "ChildPart Change"
            RevQuery = "vw_FAID_RevChildPartsGrouped"
            NoticeIDx = 17
            If FormNameX & "test" = "test" Then FormNameX = "Mass-Mod"
        Case Else
            Exit Function
    End Select
    If NoticeIDx = 16 Then GoTo ExitMe
    SQLz = ""
    SQLz = SQLz & "SELECT"
    SQLz = SQLz & " X.Action "
    SQLz = SQLz & ",X.ActionDesc "
    SQLz = SQLz & ",X.ActionDate "
    SQLz = SQLz & ",X.FieldName "
    SQLz = SQLz & ",X.OldValue "
    SQLz = SQLz & ",X.NewValue "
    SQLz = SQLz & ",X.SysUser "
    SQLz = SQLz & ",X.UnitIDStart "
    SQLz = SQLz & ",X.UnitDesc "
    SQLz = SQLz & ",X.SchID "
    SQLz = SQLz & ",X.GroupID "
    SQLz = SQLz & ",X.TableName "
    SQLz = SQLz & ",X.UnitIDEnd "
    SQLz = SQLz & "FROM " & RevQuery & " AS X "
    SQLz = SQLz & "LEFT JOIN vw_ATF_LastChange ON (X.SysUser = vw_ATF_LastChange.SysUser) "
    SQLz = SQLz & "AND (X.FieldName = vw_ATF_LastChange.FieldName) "
    SQLz = SQLz & "AND (X.GroupID = vw_ATF_LastChange.GroupID) "
    SQLz = SQLz & "WHERE X.GroupID=" & GroupIDx & " "
    If NoticeIDx = 12 Then SQLz = SQLz & "AND X.oldValue <> 'Unknown' "
    If NoticeIDx = 12 Then SQLz = SQLz & "AND X.FieldName <> 'Est. Delivery' "
    SQLz = SQLz & "AND X.ActionJustDate='" & ReportDate & "' "
    SQLz = SQLz & "AND vw_ATF_LastChange.Indx is null "

    '    Sqlz = Sqlz & "AND (ATF_LastChange.changedate <>#" & ReportDate & "# OR ATF_LastChange.Indx is null)"
    SQLz = SQLz & ";"

    If NoticeIDx = 15 Or NoticeIDx = 16 Then
        SQLz = Replace(SQLz, ",X.UnitIDStart ", "")
        SQLz = Replace(SQLz, ",X.UnitIDEnd ", "")
        SQLz = Replace(SQLz, ",X.UnitDesc ", "")
    End If
    ' 12 = Unit Vendor Change
    ' 13 = Unit Spec Change
    ' 14 = Structure / Move Units
    ' 15 = Lease Terms Change
    ' 16 = Economics Change
    ' 17 = ChildPart Change

    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then GoTo ExitMe
    RST.MoveLast
    RST.MoveFirst
    'On Error GoTo 0
    'only look for changed data and if a BLCD is SET
    If RST!Action = "Added" Or RST!ActionDesc Like "*Set*" Then
        If RST!FieldName <> "BLCD" And RST!NewValue <> "-" And NoticeIDx <> 17 And RST.RecordCount = 1 Then GoTo ExitMe
    End If

    SchIDInt = RST!SchID
    'check if Rent, Residual, Vendor or Spec if so, get the 1st unitid
    If NoticeIDx <= 14 Then UnitIDInt = CLng(RST!UnitIDStart) Else UnitIDInt = 0

    For x = 1 To RST.RecordCount
        If RST!NewValue = "Final" Then Finalized = 1
        If RST!NewValue = "Projected" Then Finalized = 2
        If RST!SysUser = FAUserName Or TestEvo = True Then
            If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then OldX = Format(RST!OldValue, "$#,#.#0") Else OldX = RST!OldValue
            If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then NewX = Format(RST!NewValue, "$#,#.#0") Else NewX = RST!NewValue
            If NoticeIDx = 14 Then
                Msgx = Msgx & "You " & SplitString(RST!ActionDesc, "from", 0) & "<Br />"
                Msgx = Msgx & "From Sch:<b> " & OldX & "</b><BR/>To Sch:<b> " & NewX & "</b><Br /><Br />"
            Else
                Msgx = Msgx & RST!ActionDesc & "<Br />"
                If NoticeIDx = 17 Then Msgx = Msgx & "item:<b> " & RST!UnitDesc & "</b><Br />"
                Msgx = Msgx & HSpaces(3) & "Prev:<b> " & OldX & "</b><Br />" & HSpaces(3) & "Curr:<b> " & NewX & "</b><Br /><Br />"
            End If
        End If
        RST.MoveNext
    Next x

    If Msgx = "" Then Exit Function

    If RST.RecordCount > 1 Then Msgx = "There were " & RST.RecordCount & " <B>" & NoticeTypeX & "</B> events<Br /><Br />" & Msgx _
        Else Msgx = "There was a <B>" & NoticeTypeX & "</B> event...<Br />" & Msgx

    Msgx = Msgx & "Please provide a REASON for the change"
    If RST.RecordCount > 1 Then Msgx = Msgx & "s"
    RST.MoveFirst
    UnitSpecX = 0
    If NoticeIDx = 14 Then Msgx = Replace(Msgx, "|", " </b>Grp<B> ", 1)
    If NoticeIDx <= 14 Then ATFIDx = Nz(DLookup("ATFID", "vw_ATFKeys", "ATFUnitSpecID=" & UnitSpecX), 0)
    'If the Group was set to 'Final', just save it and send the email
    If Finalized = 1 Then ReplyX = "Group Assignment Status set to 'Final' as per agreed terms with the Assignee"
    If Finalized = 1 Then GoTo InsertMessageLog
    'If the Group was reversed and set back to 'Projected, Request a Reason and send the email
    If Finalized = 2 Then ReplyX = ""

Prompt_For_Reason:
    If NoticeIDx = 12 And RST!FieldName = "Vendor Paid" Then
        Dim Rst2 As DAO.Recordset
        SQLx = ""
        PDFfileName = "C:\Faas\" & FAUserName & "\Temp\ATF_" & DLookup("ClientShortName", "vw_sixKeys", "Schid=" & SchIDInt) & "_Sch_" & DLookup("Schedule", "vw_sixKeys", "Schid=" & SchIDInt) & "-" & Format(Date, "mm-dd-yyyy") & ".pdf"
        SQLx = SQLx & "SELECT Count([UnitId]) AS Units, UnitInvoiceNum, Sum(UnitFinalCost) AS [AmountPaid], UnitVendorPytDate AS [DatePaid] "
        SQLx = SQLx & "From Units WHERE UnitId Between " & RST!UnitIDStart & " And " & RST!UnitIDEnd & " AND NOT UnitVendorPytDate Is Null GROUP BY UnitInvoiceNum, UnitVendorPytDate;"
        Set Rst2 = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        SQLx = ""
        Rst2.MoveLast
        Rst2.MoveFirst
        For x = 1 To Rst2.RecordCount
            InputText = InputText & "On " & CDate(Rst2!DatePaid) & ", " & Format(Rst2!AmountPaid, "$#,#.#0") & " was paid on Invoice # " & Rst2!UnitInvoiceNum & " for " & Rst2!UNITS & " units"
        Next
    End If
    If Finalized = 1 Then PDFfileName = "C:\Faas\" & FAUserName & "\Temp\ATF_" & DLookup("ClientShortName", "vw_sixKeys", "Schid=" & SchIDInt) & "_Sch_" & DLookup("Schedule", "vw_sixKeys", "Schid=" & SchIDInt) & "-" & Format(Date, "mm-dd-yyyy") & ".pdf"
    'If a change did happened, was a reason already requested, if not, then ask
    If IsReasonNeeded(GroupIDx, NoticeTypeX, CDate(ReportDate)) <> "Yes" Then
        ReplyX = IsReasonNeeded(GroupIDx, NoticeTypeX, CDate(ReportDate)) & " *previous reason"
    Else
        UserInputX = ""
        DoCmd.Beep
        ReplyX = OpenInput(Msgx, "Warning", AtlaasVersion, "Submit", InputText)
        'If Nz(ReplyX, "0") = "0" Or Len(ReplyX) < 10 Then GoTo ExitMe    ' Use ONLY For Testing !!!!
        UserStopped = False
        If ReplyX = "Stop" Then UserStopped = True
        If ReplyX = "Stop" Then GoTo InsertReason
        If Nz(ReplyX, "0") = "0" Or Len(ReplyX) < 10 Then GoTo Prompt_For_Reason
InsertReason:
        ReplyX = Replace(ReplyX, "'", "''")
        Dim NoticeDescX As String
        For x = 1 To RST.RecordCount
            If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then OldX = Format(RST!OldValue, "$#,#.#0") Else OldX = RST!OldValue
            If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then NewX = Format(RST!NewValue, "$#,#.#0") Else NewX = RST!NewValue
            If RST!ActionDesc Like "* on *" Then NoticeTypeX = SplitString(RST!ActionDesc, " on", 0)
            If RST!ActionDesc Like "* For *" Then NoticeTypeX = SplitString(RST!ActionDesc, " For", 0)
            If NoticeTypeX Like "Changed Vendor Paid*" Then NoticeTypeX = Replace(NoticeTypeX, "Changed ", "")
            NoticeDescX = RST!ActionDesc
            If NoticeDescX Like "Changed Vendor Paid*" Then NoticeDescX = Replace(NoticeDescX, "Changed ", "")
            '******* Reason(s) Inserted Here
            If TestEvo = False Then Call InsertChangeReason(RST!ActionDate, ReplyX, NoticeTypeX, RST!FieldName, RST!tableName, NoticeDescX & " (" & FormNameX & ")", Grpx, Schx, OldX, NewX, GroupIDx, SchIDInt, UnitIDInt)
            '******* Reason(s) Inserted Here
            RST.MoveNext
        Next x
    End If
    RST.MoveFirst
InsertMessageLog:
    Dim MaxINDXLog As Long
    If RST!ActionDesc Like "* on *" Then NoticeTypeX = SplitString(RST!ActionDesc, " on", 0)
    If RST!ActionDesc Like "* For *" Then NoticeTypeX = SplitString(RST!ActionDesc, " For", 0)
    If NoticeTypeX Like "Changed Vendor Paid*" Then NoticeTypeX = Replace(NoticeTypeX, "Changed ", "")
    SQLy = SQLy & "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) "
    SQLy = SQLy & "SELECT " & ClientIDint & ""
    SQLy = SQLy & ", " & SchIDInt & ""
    SQLy = SQLy & ", " & GroupIDx & ""
    SQLy = SQLy & ",#" & ReportDate & "#"
    SQLy = SQLy & ",#" & ReportDate & "#"
    SQLy = SQLy & ", " & FAUserIDx & ""
    SQLy = SQLy & ",'" & NoticeTypeX & "'"
    SQLy = SQLy & ", " & NoticeIDx & ""
    If ReplyX = "Stop" Then SQLy = SQLy & ", 1;" Else SQLy = SQLy & ", 0;"
    DoCmd.SetWarnings False
    If TestEvo = False Then DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True
    If TestEvo = False Then MaxINDXLog = DMax("INDX", "ATLAASNotifications", "UserID=" & FAUserIDx & " AND NoticeTypeID=" & NoticeIDx & "") Else MaxINDXLog = 0
    If ReplyX = "Stop" Then
        RST.close
        Exit Function
    End If

    'If a change did happened, check if an email was already set, if so, then exit
    If WasChangeEmailSent(GroupIDx, FAUserIDx, NoticeIDx, ReportDate) > 0 Then GoTo ExitMe
SendNotice:
    Dim subjectx As String
    Dim ReasonColHeader As String
    Dim ReasonRowDetail As String
    Dim Strips As Integer
    Dim Rst3   As DAO.Recordset
    NoticeDescX = NoticeTypeX
    ReasonColHeader = "Reason Provided"
    If NoticeIDx = 15 And Finalized = 1 Then
        SQLz = "SELECT SGrpID, Count(AssetID) AS Units, UnitDescShortVer, GroupOEC, BankShortName, Schedule, UnitGroup "
        SQLz = SQLz & "FROM vw_SixKeys INNER JOIN Group_OEC_Rent_LRF ON vw_SixKeys.SGrpID = Group_OEC_Rent_LRF.UnitGroupID "
        SQLz = SQLz & "GROUP BY SGrpID, UnitDescShortVer, GroupOEC, BankShortName, Schedule, UnitGroup HAVING SGrpID=" & GroupIDx & "; "
        Set Rst3 = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        Rst3.MoveLast
        Rst3.MoveFirst
        ReasonRowDetail = "<b>" & Rst3!BankShortName & "</b> Funded Schedule<b> " & Rst3!Schedule & "</b> Group<b> " & Rst3!UnitGroup & "</b> on " & Nz(DLookup("FundDate", "SchGrp", "SGrpId=" & GroupIDx), Date) & ".<br />"
        For x = 1 To Rst3.RecordCount
            ReasonRowDetail = ReasonRowDetail & "(<b>" & Rst3!UNITS & "</b>) " & Rst3!UnitDescShortVer & "<br />Totaling <b>" & Format(Rst3!GroupOEC, "$#,#.#0") & "</b><br />"
        Next
        ReplyX = ReasonRowDetail
        ReasonColHeader = "Funding Details"
        NoticeDescX = "FUNDING ALERT!<br />"
        NoticeDescX = NoticeDescX & "</span><span style='font-size:10.0pt;color:Black;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>"
        Strips = DateDiff("m", DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx & ""), DLookup("AssignmentDate", "SchGrp", "SGrpID=" & GroupIDx & ""))
        If Strips > 0 Then
           NoticeDescX = NoticeDescX & "<span style='font-size:20.0pt;color:Red;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>This Group has " & Strips & " STRIP RENTS!</span><br />"
        End If
        NoticeDescX = NoticeDescX & "<i> " & NoticeTypeX & "</i></p>"
    End If
    subjectx = ClientX
    subjectx = NoticeTypeX & "; " & ClientX & " Sch " & Schx & "|Grp " & Grpx
    'make email message
    FontSetting = "<span style='font-size:10pt;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>"
    msg1 = msg1 & "<html><body>"
    'Header Line
    msg1 = msg1 & "<p style='color:Navy;font-family:Segoe UI;font-size:20px;'>"
    msg1 = msg1 & "" & NoticeDescX & "</p>"
    'To: Line
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:13px'>To: ATF Distribution<br /><br />"
    'Regarding Line
    msg1 = msg1 & "Re. Client:  <B>" & ClientX & "</B><br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;MLNo: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Schedule: <B>" & Schx & "</B>"
    msg1 = msg1 & "&emsp;Group: <B>" & Grpx & "</B> "
    msg1 = msg1 & "&emsp;Total Units in Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><Br />"
    'Event Date Time Who Line
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Date changes made: <B> " & RST!ActionDate & "</B> "
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Changes made by: <B> " & RST!SysUser & "</B><Br />"
    'Make the Table
    msg1 = msg1 & "<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=2>"
    msg1 = msg1 & "<TR ALIGN='Center'>"
    msg1 = msg1 & FontSetting
    msg1 = msg1 & "<TH>Event</TH><TH>From</TH><TH>To</TH><TH>"
    If NoticeIDx < 15 Then msg1 = msg1 & "Current Unit Desc</TH>"
    If NoticeIDx = 15 Then msg1 = msg1 & "BLCD</TH>"
    If NoticeIDx = 17 Then msg1 = msg1 & "Child Part Desc</TH>"
    If ReplyX & "test" <> "test" Then msg1 = msg1 & "<TH>" & ReasonColHeader & "</TH>"
    msg1 = msg1 & "</TR></span>"
    For x = 1 To RST.RecordCount
        If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then OldX = Format(RST!OldValue, "$#,#.#0") Else OldX = RST!OldValue
        If NoticeIDx < 12 And RST!FieldName <> "Vendor Paid" Then NewX = Format(RST!NewValue, "$#,#.#0") Else NewX = RST!NewValue
        msg1 = msg1 & FontSetting
        msg1 = msg1 & "<TD ALIGN='Left'>" & SplitString(NoticeTypeX, " ", 0) & " " & RST!FieldName & "</TD>"
        msg1 = msg1 & "<TD ALIGN='Center'>" & OldX & "</TD>"
        msg1 = msg1 & "<TD ALIGN='Center'>" & NewX & "</TD></span>"
        msg1 = msg1 & FontSetting
        If NoticeIDx <> 15 Then msg1 = msg1 & "<TD ALIGN='Center'>" & RST!UnitDesc & "</TD>"
        If NoticeIDx = 15 Then msg1 = msg1 & "<TD ALIGN='Center'>" & DLookup("BLCD", "SchGrp", "sGrpId=" & GroupIDx & "") & "</TD>"
        If ReplyX & "test" <> "test" Then
            msg1 = msg1 & "<TD ALIGN='Left'>" & ReplyX & "</TD>"
            'If X < 2 Then msg1 = msg1 & "<TD ALIGN='Left'>" & ReplyX & "</TD>" Else msg1 = msg1 & "<TD ALIGN='Center'>-</TD>"
        End If
        msg1 = msg1 & "</span></TR>"
        RST.MoveNext
    Next x
    msg1 = msg1 & "</table>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Dim EmailToList As String
    Dim EmailCcList As String
    If PDFfileName <> "" Then Call OpenActiveATFReport(SchIDInt, True, False, True)
    EmailToList = "ATFDistribution@FleetAdvantage.com"
    If UserStopped = True Then EmailToList = "mdeaguiar@fleetadvantage.com"
    If TestEvo = True Then EmailToList = "ljohnson@fleetadvantage.com"
    EmailCcList = "ljohnson@fleetadvantage.com"

    subjectx = Replace(subjectx, "'", "''") & " - ATLAAS Core Message"
    If UserStopped = True Then subjectx = subjectx & " USER STOPPED!"
    msg1 = Replace(msg1, "'", "''")
    If TestEvo = False Then
        If FAUserName = "ytj" Then
            Call SendServerEmail(EmailToList, EmailCcList, subjectx, msg1, "yjackson@fleetadvantage.com", PDFfileName)
        Else
            Call SendServerEmail(EmailToList, EmailCcList, subjectx, msg1, FAUserName & "@fleetadvantage.com", PDFfileName)
        End If
    Else
        'Call SendEmailMessage(EmailToList, "ATLAAS Core Message: " & subjectx & "", msg1, False, True, False, EmailCcList)
        Call SendServerEmail(EmailToList, EmailCcList, subjectx, msg1, FAUserName & "@fleetadvantage.com")
    End If
    'Mark Email Sent in the notifaction log
    DoCmd.SetWarnings False
    If TestEvo = False Then DoCmd.RunSQL ("Update ATLAASNotifications Set NoticeSent=1 Where INDX=" & MaxINDXLog & ";")
    DoCmd.SetWarnings True
    Exit Function
ExitMe:
    On Error GoTo FixMe
    RST.close
    Set RST = Nothing
FixMe:
    On Error GoTo 0
End Function


Public Function SetATFDataTable(FormNameX As String)
    Forms!ATFTreeCommand(FormNameX).Form.DatasheetFontName = "Segoe UI"
    Forms!ATFTreeCommand(FormNameX).Form.DatasheetFontHeight = 9
    Forms!ATFTreeCommand(FormNameX).Form.DatasheetFontWeight = 400
End Function

Public Function OpenScheduleControl(UnitIDx As Long)
    DoCmd.OpenForm "ScheduleControl"
    Forms!ScheduleControl.Form.SearchAssetID2.SetFocus
    Forms!ScheduleControl.Form.SearchAssetID2.Text = UnitIDx
    Forms!ScheduleControl.Form.UnitLookUp = UnitIDx
    Form_ScheduleControl.Remotelookup
    DoCmd.SelectObject acForm, "ScheduleControl"
End Function
Public Function WasChangeEmailSent(GroupIDy As Long, UserIDy As Integer, NoticeIDy As Integer, ReportDatey As String) As Integer
    WasChangeEmailSent = DCount("Indx", "ATLAASNotifications", "NoticeSent=1 AND GroupId=" & GroupIDy & " AND UserID=" & UserIDy & " AND  NoticeTypeId=" & NoticeIDy & " AND NoticeDate= #" & ReportDatey & "#")
    'AND UserID=" & UserIDy & "
End Function
Public Function IsReasonNeeded(GroupIDx As Long, ReasonType As String, Optional ReportDate As String) As String
    IsReasonNeeded = "Yes"
    Dim SQLz   As String
    If ReportDate & "test" = "test" Then ReportDate = Date
    SQLz = ""
    SQLz = "SELECT UserID, GroupID, NoticeType, NoticeDate, NoticeSent From ATLAASNotifications " & _
           "WHERE UserID = " & FAUserIDx & " AND  GroupID=" & GroupIDx & " AND NoticeType='" & ReasonType & "' AND NoticeDate = #" & ReportDate & "#;"

    Dim RST    As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo ExitMe
    If RST.RecordCount >= 1 Then IsReasonNeeded = RST!noticeType
ExitMe:
    RST.close
    Set RST = Nothing
End Function
Public Function InsertChangeReason(ActionDate As Date, Reason, FieldGroup, FieldName, tableName, _
                                   Action, groupNo, SchNo, OldValue, NewValue As String, _
                                   groupID, SchID, UnitID As Long)
' ATFID, ATFSchSpecID, ATFGroupSpecID, ATFUnitSpecID,
    Dim SQLy   As String
    SQLy = ""
    SQLy = SQLy & "INSERT INTO ATF_ChangeReasons ("
    SQLy = SQLy & "ChangeDate"
    SQLy = SQLy & ",Reason"
    SQLy = SQLy & ",FieldGroup"
    SQLy = SQLy & ",FieldName"
    SQLy = SQLy & ",SysUser"
    SQLy = SQLy & ",TableName"
    SQLy = SQLy & ",Action"
    SQLy = SQLy & ",GroupNo"
    SQLy = SQLy & ",GroupID "
    SQLy = SQLy & ",SchNo"
    SQLy = SQLy & ",SchID "
    SQLy = SQLy & ",OldValue "
    SQLy = SQLy & ",NewValue "
    SQLy = SQLy & ",UnitID "
    SQLy = SQLy & ",RevID "
    SQLy = SQLy & ") Select "
    SQLy = SQLy & " '" & ActionDate & "'"
    SQLy = SQLy & ",'" & Action & vbNewLine & Reason & "'"
    SQLy = SQLy & ",'" & FieldGroup & "'"
    SQLy = SQLy & ",'" & FieldName & "'"
    SQLy = SQLy & ",'" & FAUserName & "'"
    SQLy = SQLy & ",'" & tableName & "'"
    SQLy = SQLy & ",'" & Replace(Action, "'", "''") & "'"
    SQLy = SQLy & ",'" & Replace(groupNo, "'", "''") & "'"
    SQLy = SQLy & "," & groupID & ""
    SQLy = SQLy & ",'" & Replace(SchNo, "'", "''") & "'"
    SQLy = SQLy & "," & SchID & ""
    SQLy = SQLy & ",'" & OldValue & "'"
    SQLy = SQLy & ",'" & NewValue & "'"
    SQLy = SQLy & "," & UnitID & ""
    SQLy = SQLy & "," & NextRevID(SchID) & ""
    SQLy = SQLy & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True
End Function
Public Function MakeATFRevision(SchIDx As Long, ReasonX As String, CheckRevMadeToday As Boolean, ShowDialog As Boolean)
    Dim RST    As DAO.Recordset
    Dim qdf    As DAO.QueryDef
    Dim SQLx   As String
    Dim i, RevNum As Integer
    SQLx = "Select * From vw_ATF_MakeScheduleSnapShot"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    SQLx = ""
    RST.MoveLast
    RST.MoveFirst
    RevNum = CInt(Nz(DMax("RevisionNum", "ATFScheduleRevisions", "SchID=" & SchIDx), 0)) + 1
    If CDate(Nz(DMax("RevisionDate", "ATFScheduleRevisions", "SchID =" & SchIDx), "1/1/1900")) = Date And CheckRevMadeToday = True _
       And FAUserName = DLookup("SysUser", "ATFScheduleRevisions", "SchID=" & SchIDx & "AND RevisionNum=" & RevNum - 1 & "") Then
        DoCmd.Beep
        Select Case AsktoReplace
        Case 1
            GoTo PostRevision
        Case 2
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Delete * From ATFScheduleRevisions Where SchID= " & SchIDx & " AND RevisionDate = #" & Date & "#")
            DoCmd.SetWarnings True
            If RevNum > 1 Then RevNum = RevNum - 1
        Case 3
            Call Dialog.RichBox(Prompt:="<font Size=""2"">Revision Canceled</Font><Br />No Updates to the revision table were made", Buttons:=vbOKOnly + vbExclamation, Title:=AtlaasVersion)
            GoTo FixMe
        End Select
    End If
PostRevision:
 SQLx = "SELECT " & RevNum & " As RevisionNum, #" & Date & "# as RevisionDate, '" & FAUserName & "' as SysUser, '" & ReasonX & "' as RevisionNote, "
    SQLx = SQLx & RST!SelectPart1
    'Get Fields
    For i = 1 To RST.RecordCount
        SQLx = SQLx & RST!SelectPart2
        RST.MoveNext
    Next i
    RST.MoveFirst
    'Get Tables
    SQLx = SQLx & RST!FromPart
    SQLx = SQLx & vbNewLine
    SQLx = SQLx & " WHERE Schedule.SchID=" & SchIDx & " "
    SQLx = SQLx & vbNewLine
    'Get Group by
    SQLx = SQLx & RST!GroupPart1
    SQLx = SQLx & vbNewLine
    For i = 1 To RST.RecordCount
        If RST!GroupPart2 <> "-" Then SQLx = SQLx & RST!GroupPart2
        RST.MoveNext
    Next i
    RST.MoveFirst
    SQLx = SQLx & RST!HavingPart & ";"
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "ATFScheduleRevisionsTemp"
    On Error GoTo 0
    Set qdf = CurrentDb.CreateQueryDef("ATFScheduleRevisionsTemp", SQLx)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "ATFScheduleRevisions_Append"
    DoCmd.SetWarnings True
    If ShowDialog = True Then Call Dialog.RichBox(Prompt:="<font Size=""3"">Revision " & RevNum & " Completed</Font>", Buttons:=vbOKOnly + vbInformation, Title:=AtlaasVersion, AllowCopyToClipboard:=False, AllowSaveToFile:=False, ButtonDelay:=0)
FixMe:
    RST.close
    Set RST = Nothing
End Function
Public Function AsktoReplace() As Integer
    Dim dr     As VbMsgBoxResultEx
    dr = Dialog.RichBox(Prompt:="<font Size=""3.5"">Replace or Add?</font><Br />" & FAUserFullName & ", you made a Schedule Revision eariler today…<Br />Do you want to<B><font color=""#FF0000""> Replace it </font></B>or<B> Add it </B> as another revision?", _
                        Buttons:=vbCustom + vbQuestion, _
                        Title:=AtlaasVersion, _
                        AllowBeep:=True, _
                        LabelButton1:="Add", _
                        LabelButton2:="Replace", _
                        LabelButton3:="Cancel Revision")
    AsktoReplace = (dr - 100)
    'AllowCopyToClipboard:=False, _
     AllowSaveToFile:=False, _
     ButtonDelay:=0,
End Function

Public Function InsertSchedule(MLIDi As Integer, SourceSchIDi As Integer, SchNoX As String, SchShortDescX As String)
    Dim SQLz   As String
    SQLz = SQLz & "INSERT INTO [Schedule] "
    SQLz = SQLz & "([SchMLRef],[SchNo],[DateSchAdded],[SchShortDesc],[SchAlias]"
    SQLz = SQLz & ",[BankId],[SchFSL],[SchOut],[SchIn],[Schsigned],[SchSigner],[SchSigner2],[SchSigner3]"
    SQLz = SQLz & ",[SchAssDt],[SchTitleAssignee],[SchTitleToTrust],[SchTitleHolder],[SchFAID],[SCHINSAMT],[SCHINREC],[SchInAccpt],[SCHINSEXP],[SchPhysInsRec],[SchPInsAccpt]"
    SQLz = SQLz & ",[SchPinsExp],[SchTitltoFA],[SchTitletoClient],[LienHolder],[LienHolderNm],[FSLSCH],[SchPDInv],[SchCurrRent],[SchCRExp],[SchPDInvDt],[SchPDBillDt]"
    SQLz = SQLz & ",[SchBnkClNO],[SchBillAdd],[POVendor],[POSentDt],[PO2Vendor],[PO2SentDt],[PO3Vendor],[PO3SentDt],[PilotRev],[PilotRevdt],[PeopleNet],[PNMOChg],[BLCDACDT]"
    SQLz = SQLz & ",[PNUnit],[PNInstall],[UpFrontComm],[FundCommPmt],[CPMModel],[IntLease],[InvFld1desc],[InvFld1],[InvFld2desc],[InvFld2],[ATFrev],[Revdt],[SchReviewed]"
    SQLz = SQLz & ",[SchRver],[SCH-A],[SchSLVRec],[SchedDt],[SchAH],[SchFABill],[SchExSyn],[SchFEcoRec],[SchFEcoA],[SchFPkgO],[SchRALO],[SchRALB],[SchRALSigner]"
    SQLz = SQLz & ",[SchCoNo],[SCHNo1],[SchIntCommbase],[FSLType],[SubLeasee],[FAVendorNumber],[FAPOnumber],[ApprovalRequest],[ApprovalReply],[UseSpecialRemitSch]"
    SQLz = SQLz & ",[SpecialRemitAddrSch],[ATFSchSpecID],[ATFID])"
    SQLz = SQLz & "Select "
    SQLz = SQLz & MLIDi & " AS [SchMLRef]"
    SQLz = SQLz & ",'" & SchNoX & "' As [SchNo]"
    SQLz = SQLz & ",#" & Date & "# As [DateSchAdded]"
    SQLz = SQLz & ",'" & SchShortDescX & "' As [SchShortDesc]"
    SQLz = SQLz & ",'' as [SchAlias]"
    SQLz = SQLz & ",[BankId],[SchFSL],[SchOut],[SchIn],[Schsigned],[SchSigner],[SchSigner2],[SchSigner3]"
    SQLz = SQLz & ",[SchAssDt],[SchTitleAssignee],[SchTitleToTrust],[SchTitleHolder],[SchFAID],[SCHINSAMT],[SCHINREC],[SchInAccpt],[SCHINSEXP],[SchPhysInsRec],[SchPInsAccpt]"
    SQLz = SQLz & ",[SchPinsExp],[SchTitltoFA],[SchTitletoClient],[LienHolder],[LienHolderNm],[FSLSCH],[SchPDInv],[SchCurrRent],[SchCRExp],[SchPDInvDt],[SchPDBillDt]"
    SQLz = SQLz & ",[SchBnkClNO],[SchBillAdd],[POVendor],[POSentDt],[PO2Vendor],[PO2SentDt],[PO3Vendor],[PO3SentDt],[PilotRev],[PilotRevdt],[PeopleNet],[PNMOChg],[BLCDACDT]"
    SQLz = SQLz & ",[PNUnit],[PNInstall],[UpFrontComm],[FundCommPmt],[CPMModel],[IntLease],[InvFld1desc],[InvFld1],[InvFld2desc],[InvFld2],[ATFrev],[Revdt],[SchReviewed]"
    SQLz = SQLz & ",[SchRver],[SCH-A],[SchSLVRec],[SchedDt],[SchAH],[SchFABill],[SchExSyn],[SchFEcoRec],[SchFEcoA],[SchFPkgO],[SchRALO],[SchRALB],[SchRALSigner]"
    SQLz = SQLz & ",[SchCoNo],[SCHNo1],[SchIntCommbase],[FSLType],[SubLeasee],[FAVendorNumber],[FAPOnumber],[ApprovalRequest],[ApprovalReply],[UseSpecialRemitSch]"
    SQLz = SQLz & ",[SpecialRemitAddrSch],[ATFSchSpecID],[ATFID]"
    SQLz = SQLz & " From Schedule WHERE SchID = " & SourceSchIDi & ""
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Public Function InsertSchedule_Profile(MLIDi As Integer, ByRef ATFIDi As Integer, SchNoX As String, FAIDx As String)
    Dim SQLz   As String
    SQLz = ""
    SQLz = SQLz & "INSERT INTO ATFScheduleSpec ("
    SQLz = SQLz & "ATFSchName"
    SQLz = SQLz & ", SchFAID"
    SQLz = SQLz & ", ATFID"
    SQLz = SQLz & ", MLID"
    SQLz = SQLz & ", FSLType"
    SQLz = SQLz & ", SchSpecInserted"
    SQLz = SQLz & ", NumGroupsInSch"
    SQLz = SQLz & ", ATFSchSpecDesc"
    SQLz = SQLz & ", SchDescription"
    SQLz = SQLz & ") Values ( "
    SQLz = SQLz & " '" & SchNoX & "' "
    SQLz = SQLz & ",'" & FAIDx & "' "
    SQLz = SQLz & ", " & ATFIDi & ""
    SQLz = SQLz & ", " & MLIDi & ""
    SQLz = SQLz & ", 'Yes'"
    SQLz = SQLz & ", #" & Date & "#"
    SQLz = SQLz & ", 1"
    SQLz = SQLz & ", '" & SchNoX & "' "
    SQLz = SQLz & ", '" & DLookup("ATFDesc", "ATF", "ATFID =" & ATFIDi) & "' "
    SQLz = SQLz & ")"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Public Function InsertGroup(FAIDx, NewGroupNo As String, SourceGroupID, SchIDx As Long, Optional FormNmx As String, Optional CloneIt As Boolean)
    Dim SQLz   As String
    SQLz = "INSERT INTO schgrp ( "
    SQLz = SQLz & "SGrpNo"
    SQLz = SQLz & ", SGrpSchID"
    SQLz = SQLz & ", SGrpFAID"
    SQLz = SQLz & ", GroupAddedDate"
    SQLz = SQLz & ", SgrpDesc"
    SQLz = SQLz & ",Assigned_Indicator"
    SQLz = SQLz & ", UnitsInGroup"
    SQLz = SQLz & ", OrigEquipCost"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ", ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", LeaseExtOption1Term"
    SQLz = SQLz & ", LeaseExtOption1Rent"
    SQLz = SQLz & ", MTMRent"
    SQLz = SQLz & ", LeaseTermPrimary"
    SQLz = SQLz & ", LeaseTermNotice"
    SQLz = SQLz & ", LeasePymtFreq"
    SQLz = SQLz & ", LeasePaidAdvArr"
    SQLz = SQLz & ", LeasePymtForGroup"
    SQLz = SQLz & ", LRFGroup"
    SQLz = SQLz & ", SGRPPLB"
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", SgrpCommInd"
    SQLz = SQLz & ", PerDiemStartsOn"

    'above is min required to deploy a group
    'below is desired min to deploy a group
    If Nz(SourceGroupID, 0) > 0 Then
        'If there is a source groupid then add extended fields below
        SQLz = SQLz & ", SGrpSalesIndex"
        SQLz = SQLz & ", SGrpSalesIndDt"
        SQLz = SQLz & ", SgrpSalesIndRt"
        SQLz = SQLz & ", SGprSalesIndYr"
        SQLz = SQLz & ", SgrpSalesSpread"
        SQLz = SQLz & ", SgrpSalesYield"
        SQLz = SQLz & ", SGrpBonus"
        SQLz = SQLz & ", BPChgLO"
        SQLz = SQLz & ", RebateUnitProjected"
        SQLz = SQLz & ", AllowedHoursPer"
        SQLz = SQLz & ", SGprAllow1"
        SQLz = SQLz & ", SGrpov1"
        SQLz = SQLz & ", SGprAllow2"
        SQLz = SQLz & ", SGrpov2"
        SQLz = SQLz & ", SGprAllow4"
        SQLz = SQLz & ", SGrpov4"
        SQLz = SQLz & ", SGprAllow5"
        SQLz = SQLz & ", SGrpov5"
        SQLz = SQLz & ", SGrpFSLPmt"
        SQLz = SQLz & ", ScheduleGroupPaymentTerm"
        SQLz = SQLz & ", SGrpRSLTPmt"
    End If
    'Data to use part
    SQLz = SQLz & ") SELECT "
    SQLz = SQLz & "'" & NewGroupNo & "'"
    SQLz = SQLz & "," & SchIDx & ""
    SQLz = SQLz & ", '" & FAIDx & "'"
    SQLz = SQLz & ", #" & Date & "# "
    'Is the the request is coming from a form with data
    If FormNmx <> "" Then
        SQLz = SQLz & ",'" & Forms(FormNmx).SgrpDesc & "'"
        SQLz = SQLz & ", 'P'"
        SQLz = SQLz & ", " & Forms(FormNmx).UnitsInGroup & ""
        SQLz = SQLz & ", " & Forms(FormNmx).OrigEquipCost & ""
        SQLz = SQLz & ", " & Forms(FormNmx).ExchangeITGroup & ""
        SQLz = SQLz & ", '" & Forms(FormNmx).ExchangeITType & "'"
        SQLz = SQLz & ", " & Forms(FormNmx).ExchangeITMonth & ""
        SQLz = SQLz & ", " & Forms(FormNmx).LeaseExtOption1Term & ""
        SQLz = SQLz & ", " & Forms(FormNmx).LeaseExtOption1Rent & ""
        SQLz = SQLz & ", " & Forms(FormNmx).MTMRent & ""
        SQLz = SQLz & ", " & Forms(FormNmx).LeaseTermPrimary & ""
        SQLz = SQLz & ", " & Forms(FormNmx).LeaseTermNotice & ""
        SQLz = SQLz & ", '" & Forms(FormNmx).LeasePymtFreq & "'"
        SQLz = SQLz & ", '" & Forms(FormNmx).LeasePaidAdvArr & "'"
        SQLz = SQLz & ", " & Forms(FormNmx).LeasePymtForGroup & ""
        SQLz = SQLz & ", " & Forms(FormNmx).LRFGroup & ""
        SQLz = SQLz & ", '" & Forms(FormNmx).SGRPPLB & "'"
        SQLz = SQLz & ", '" & Forms(FormNmx).LeaseType & "'"
        SQLz = SQLz & ", " & Nz(Forms(FormNmx).SgrpCommInd, 10) & ""
        SQLz = SQLz & ",  " & Nz(Forms(FormNmx).PerDiemStartsOn, 10) & ""
    Else
        'If no form then use defaults instead
        SQLz = SQLz & ", 'Added in by ATLAAS Core Automation'"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", ''"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 72"
        SQLz = SQLz & ", 3"
        SQLz = SQLz & ", 'M'"
        SQLz = SQLz & ", 'Adv'"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ", 'No'"
        SQLz = SQLz & ", 'FMV'"
        SQLz = SQLz & ", 1"
        SQLz = SQLz & ", 2"
    End If
    If Nz(SourceGroupID, 0) > 0 Then
        'If there is a source groupid then add extended data below
        SQLz = SQLz & ", SGrpSalesIndex"
        SQLz = SQLz & ", SGrpSalesIndDt"
        SQLz = SQLz & ", SgrpSalesIndRt"
        SQLz = SQLz & ", SGprSalesIndYr"
        SQLz = SQLz & ", SgrpSalesSpread"
        SQLz = SQLz & ", SgrpSalesYield"
        SQLz = SQLz & ", SGrpBonus"
        SQLz = SQLz & ", BPChgLO"
        SQLz = SQLz & ", RebateUnitProjected"
        SQLz = SQLz & ", AllowedHoursPer"
        SQLz = SQLz & ", SGprAllow1"
        SQLz = SQLz & ", SGrpov1"
        SQLz = SQLz & ", SGprAllow2"
        SQLz = SQLz & ", SGrpov2"
        SQLz = SQLz & ", SGprAllow4"
        SQLz = SQLz & ", SGrpov4"
        SQLz = SQLz & ", SGprAllow5"
        SQLz = SQLz & ", SGrpov5"
        SQLz = SQLz & ", SGrpFSLPmt"
        SQLz = SQLz & ", ScheduleGroupPaymentTerm"
        SQLz = SQLz & ", SGrpRSLTPmt"
        SQLz = SQLz & " FROM SchGrp"
        SQLz = SQLz & " WHERE SGrpId=" & Nz(SourceGroupID, 0) & ""
    End If
    SQLz = SQLz & ";"
    DoCmd.SetWarnings False
    If CloneIt = True Then
        DoCmd.OpenQuery ("GroupAdd_AppendQuery")
    Else
        DoCmd.RunSQL (SQLz)
    End If

    DoCmd.SetWarnings True
End Function
Public Function InsertGroup_Profile(ATFSchSpecIDi As Integer, NumOfUnitsi As Integer, GroupNoX As String)
    Dim SQLz
    SQLz = SQLz & "INSERT INTO ATFGroupSpec ( "
    SQLz = SQLz & "ATFID"
    SQLz = SQLz & ", ATFSchSpecID"
    SQLz = SQLz & ", ATFGroupSpecDesc"
    SQLz = SQLz & ", UnitSpecsInGroup"
    SQLz = SQLz & ", UnitsInGroup"
    SQLz = SQLz & ", GrpFAID"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ", ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", ExtendITOptionTerm"
    SQLz = SQLz & ", PrimaryLeaseTerm"
    SQLz = SQLz & ", PrimaryLeaseNotice"
    SQLz = SQLz & ", LeasePymtFreq"
    SQLz = SQLz & ", LeasePaidAdvArr"
    SQLz = SQLz & ", SwapIndex"
    SQLz = SQLz & ", SwapDate"
    SQLz = SQLz & ", SwapRate"
    SQLz = SQLz & ", SwapTermYears"
    SQLz = SQLz & ", SwapYield"
    SQLz = SQLz & ", SwapBonus"
    SQLz = SQLz & ", SwapAdjustment"
    SQLz = SQLz & ", ResidualAvg"
    SQLz = SQLz & ", UATPerUnit"
    SQLz = SQLz & ", PLB"
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", BLCDBasisID"
    SQLz = SQLz & ", PerDiemBasisID"
    SQLz = SQLz & ", MPYAllowance"
    SQLz = SQLz & ", HPYAllowance"
    SQLz = SQLz & ", AllowanceMethod"
    SQLz = SQLz & ", CPMileThreshold"
    SQLz = SQLz & ", CPHourThreshold )"
    SQLz = SQLz & " Select TOP 1 "
    SQLz = SQLz & "ATFID"
    SQLz = SQLz & ", ATFSchSpecID"
    SQLz = SQLz & ", 'Group " & GroupNoX & "'"
    SQLz = SQLz & ", UnitSpecsInGroup"
    SQLz = SQLz & ", " & NumOfUnitsi & ""
    SQLz = SQLz & ", GrpFAID"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ", ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", ExtendITOptionTerm"
    SQLz = SQLz & ", PrimaryLeaseTerm"
    SQLz = SQLz & ", PrimaryLeaseNotice"
    SQLz = SQLz & ", LeasePymtFreq"
    SQLz = SQLz & ", LeasePaidAdvArr"
    SQLz = SQLz & ", SwapIndex"
    SQLz = SQLz & ", SwapDate"
    SQLz = SQLz & ", SwapRate"
    SQLz = SQLz & ", SwapTermYears"
    SQLz = SQLz & ", SwapYield"
    SQLz = SQLz & ", SwapBonus"
    SQLz = SQLz & ", SwapAdjustment"
    SQLz = SQLz & ", ResidualAvg"
    SQLz = SQLz & ", UATPerUnit"
    SQLz = SQLz & ", PLB"
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", BLCDBasisID"
    SQLz = SQLz & ", PerDiemBasisID"
    SQLz = SQLz & ", MPYAllowance"
    SQLz = SQLz & ", HPYAllowance"
    SQLz = SQLz & ", AllowanceMethod"
    SQLz = SQLz & ", CPMileThreshold"
    SQLz = SQLz & ", CPHourThreshold "
    SQLz = SQLz & "From ATFGroupSpec WHERE ATFSchSpecID=" & ATFSchSpecIDi
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Public Function SetUnitLevelHash(ATFx, GrpIdx, ATFGroupSpecX As Integer)
    Dim SQLz   As String
    SQLz = "INSERT INTO ATF_UnitLevelHash ( ATFID, ATFUnitSpecID, UnitID, ATFHash_Rent, ATFHash_Spec, ATFHash_Vendor, ATFHash_RentLastDate, ATFHash_SpecLastDate, ATFHash_VendorLastDate, GroupID, ATFGroupSpecID ) " & _
           "SELECT ATFID, ATFUnitSpecID, UnitID, RentHash_New, SpecHash_New, VendorHash_New, Now() AS Expr1, Now() AS Expr2, Now() AS Expr3," & GrpIdx & ", " & ATFGroupSpecX & " " & _
           "FROM vw_ATF_UnitLevelHashCheck WHERE ATFID=" & ATFx & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function

Public Function OpenActiveATFReport_OLD(SchIDx As Long)
    DoCmd.close acReport, "ATF_Report_Active"
    DoCmd.OpenReport "ATF_Report_Active", acViewPreview, , "vw_ATF_Active_Profile.SchID=" & SchIDx & ""
End Function
Public Function DeleteExtraATFReports()
    On Error Resume Next
    Dim Rptx   As Object
    For Each Rptx In Application.CurrentProject.AllReports
        If Rptx.Name Like "ATF_Report_Active_*" Then DoCmd.DeleteObject acReport, Rptx.Name
    Next Rptx
    On Error GoTo 0
End Function
Public Function OpenActiveATFReport(SchIDx As Long, Optional MakePDF As Boolean, Optional ImportNotice As Boolean, Optional JustPDF As Boolean)
    On Error Resume Next
    Dim Rptx   As Object
    Dim NewReport As String
    Dim xi     As Integer
    Dim SQLz   As String

    'Check if any units are set for upFront Tax, if so tell the user.
    SQLz = "SELECT vw_SixKeys.SchID"
    SQLz = SQLz & ",Count(vw_SixKeys.AssetID) AS upFrontUnits "
    SQLz = SQLz & "FROM (vw_SixKeys INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID)"
    SQLz = SQLz & " WHERE SchID = " & SchIDx & " And isUpFront = 1 "
    SQLz = SQLz & " GROUP BY vw_SixKeys.SchID;"

    Dim RST    As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 And RST!upFrontUnits = 1 Then Call Dialog.RichBox(Prompt:="<font Size=""3"">" & RST!upFrontUnits & " Unit is set for<b> UpFront</b> Tax</Font>", Buttons:=vbOKOnly + vbInformation, Title:=AtlaasVersion, AllowCopyToClipboard:=False, AllowSaveToFile:=False, ButtonDelay:=0)
    If RST.RecordCount > 0 And RST!upFrontUnits > 1 Then Call Dialog.RichBox(Prompt:="<font Size=""3"">" & RST!upFrontUnits & " Units are set for<b> UpFront</b> Tax</Font>", Buttons:=vbOKOnly + vbInformation, Title:=AtlaasVersion, AllowCopyToClipboard:=False, AllowSaveToFile:=False, ButtonDelay:=0)
    RST.close
    Set RST = Nothing

'Make  the ATF Report
    Call MessageUser("Open", "Making ATF Report", "Defining Queries...")
SetQueryDef:
    CurrentDb.QueryDefs("ATF_ReportActiveQ").SQL = "Select * From ATF_ReportActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_EquipmentActiveQ").SQL = "Select * From ATF_Report_EquipmentActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_ChildPartsActiveQ").SQL = "Select * From ATF_Report_ChildPartsActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_AllowancesActiveQ").SQL = "Select * From ATF_Report_AllowancesActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_LeaseActiveQ").SQL = "Select * From ATF_Report_LeaseActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_SchNotesActiveQ").SQL = "Select * From ATF_Report_SchNotesActive Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_FAID_ContactsMLSigners").SQL = "Select * From ATF_FAID_ContactsMLSignersBase Where SchID=" & SchIDx & ";"
    CurrentDb.QueryDefs("ATF_Report_BillingActiveQ").SQL = "Select * From ATF_Report_BillingActive Where SchID=" & SchIDx & ";"
    'ATF_ReportActiveQ
    Call MessageUser("Update", "Making ATF Report", "Setting Clearing Previous Reports, working...")
    For Each Rptx In Application.CurrentProject.AllReports
        If Rptx.Name Like "ATF_Report_Active_*" Then DoCmd.DeleteObject acReport, Rptx.Name
    Next Rptx
    For Each Rptx In Application.CurrentProject.AllReports
        If Rptx.Name Like "ATF_Report_Active_*" Then xi = xi + 1
    '    Call MessageUser("Update", "Making ATF Report-" & xi, "Setting Clearing Previous Reports, working...")
    Next Rptx
    On Error GoTo OpenActiveATFReport_Error
SetReportNumber:
    Call MessageUser("Update", "Making ATF Report", "Creating 'New' Reports, Just a bit more...")
    xi = xi + 1
    NewReport = CStr("ATF_Report_Active_" & xi)
    For Each Rptx In Application.CurrentProject.AllReports
        If Rptx.Name = NewReport Then GoTo SetReportNumber
    Next Rptx
    DoCmd.CopyObject , NewReport, acReport, "ATF_Report_Active"
    If JustPDF <> True Then DoCmd.OpenReport NewReport, acViewPreview
    ' DoCmd.OpenReport NewReport, acViewPreview, , "vw_ATF_Active_Profile.SchID=" & SchIDx & ""
    If MakePDF = True Then
        Call MessageUser("Update", "Making ATF Report", "Creating 'New' PDF to attach...")
        Dim strFileName As String
        strFileName = "C:\Faas\" & FAUserName & "\Temp\ATF_" & DLookup("ClientShortName", "vw_sixKeys", "Schid=" & SchIDx) & "_Sch_" & DLookup("Schedule", "vw_sixKeys", "Schid=" & SchIDx) & "-" & Format(Date, "mm-dd-yy") & ".pdf"
        DoCmd.OutputTo acOutputReport, NewReport, acFormatPDF, strFileName, False
    End If
    If MakePDF = True And ImportNotice = True Then
        Call MessageUser("Update", "Making ATF Report", "Creating Email, Almost done...")
        Call SendATFImportNotice(SchIDx, strFileName)
    End If
    Call MessageUser("Close")
    On Error GoTo 0
    Exit Function
OpenActiveATFReport_Error:
    Call MessageUser("Close")
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenActiveATFReport of Module ATF_Functions"
End Function
Public Function NextRevID(SchIDx) As Integer
    NextRevID = CInt(Nz(DMax("RevID", "ATF_ChangeReasons", "SchID=" & SchIDx & ""), 0)) + 1
End Function

Public Function SendATFImportNotice(SchIDx As Long, FileToAttachment As String)
    Dim SQLy   As String
    Dim msg1, FontSetting As String
InsertMessageLog:
    SQLy = SQLy & "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) "
    SQLy = SQLy & "SELECT " & DLookup("ClientID", "vw_SixKeys", "SchID=" & SchIDx) & ""
    SQLy = SQLy & ", " & SchIDx & ""
    SQLy = SQLy & ", " & 0 & ""
    SQLy = SQLy & ",#" & Date & "#"
    SQLy = SQLy & ",#" & Date & "#"
    SQLy = SQLy & ", " & FAUserIDx & ""
    SQLy = SQLy & ",'ATF Imported Notice'"
    SQLy = SQLy & ",16"
    SQLy = SQLy & ", 1;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True
SendNotice:
    'make email message
    FontSetting = "<span style='font-size:10pt;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>"
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>The LPAF for FAID: " & DLookup("FAID", "vw_SixKeys", "SchID=" & SchIDx) & " was imported from Sales Force</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:13px'>To: ATF Distribution<br /><br />"
    msg1 = msg1 & "Re. Client:  <B>" & DLookup("CompanyName", "vw_SixKeys", "SchID=" & SchIDx) & "</B><br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;MLNo: <B>" & DLookup("MLNo", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Schedule Number: <B>" & DLookup("Schedule", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Number of Groups: <B>" & DCount("SGrpSchID", "SchGrp", "SGrpSchId=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Total Units in ATF: <B>" & DCount("UnitID", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Date of Import: <B> " & DLookup("ChangeDate", "ATF_ChangeReasons", "Action ='Schedule Import' AND Schid=" & SchIDx & "") & "</B><Br /> "
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Import ran by: <B> " & FAUserName & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Please see the attached PDF"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Dim EmailToList As String
    'EmailToList = "kzeigler@FleetAdvantage.com;pachille@FleetAdvantage.com;bmcmahon@FleetAdvantage.com;JGriffin@FleetAdvantage.com;rwasserman@FleetAdvantage.com"
    EmailToList = "ATFDistribution@FleetAdvantage.com"
    Call SendEmailMessage(EmailToList, "ATLAAS Core Message: ATF Import", msg1, True, True, False, "ljohnson@fleetadvantage.com", , , FileToAttachment)
End Function
Public Function CleanRevComments()
    Dim SQLz, ReplaceX As String
    Dim Tempx  As String
    Dim ReplaceWith As String
    Dim xi, Zi As Integer
    For xi = 1 To 26
        If xi = 1 Then ReplaceX = "size=6"
        If xi = 1 Then ReplaceWith = "size=2"
        If xi = 2 Then ReplaceX = "size=5"
        If xi = 2 Then ReplaceWith = "size=2"
        If xi = 3 Then ReplaceX = "size=4"
        If xi = 3 Then ReplaceWith = "size=2"
        If xi = 4 Then ReplaceX = "size=3"
        If xi = 4 Then ReplaceWith = "size=2"
        If xi = 5 Then ReplaceX = "<div><font size=3>&nbsp;&nbsp;</font></div>"
        If xi = 5 Then ReplaceWith = ""
        If xi = 6 Then ReplaceX = "<div><font size=3>&nbsp;</font></div>"
        If xi = 6 Then ReplaceWith = ""
        If xi = 7 Then ReplaceX = "<div>&nbsp;</div>"
        If xi = 7 Then ReplaceWith = ""
        If xi = 8 Then ReplaceX = "<div><font size=2 color=red><strong>&nbsp;&nbsp;</B></font></div>"
        If xi = 8 Then ReplaceWith = ""
        If xi = 9 Then ReplaceX = "<div><font size=2 color=red><B>&nbsp;</B></font></div>"
        If xi = 9 Then ReplaceWith = ""
        If xi = 10 Then ReplaceX = "<div><font size=2>&nbsp;</font></div>"
        If xi = 10 Then ReplaceWith = ""
        If xi > 10 Then
            Tempx = ""
            For Zi = (29 - xi) To 3 Step -1
                Tempx = Tempx & "&nbsp;"
            Next Zi
            ReplaceX = Tempx
            ReplaceWith = ""
        End If
        If DCount("RevId", "ATFRevT", "Comment Like '*" & ReplaceX & "*'") > 0 Then
            Call DoSql(ReplaceX, ReplaceWith)
            xi = xi - 1
        Else
            xi = xi
        End If
        If DCount("INDX", "ATF_ChangeReasons", "Reason Like '*" & ReplaceX & "*'") > 0 Then
            Call DoSqlReason(ReplaceX, ReplaceWith)
            xi = xi - 1
        Else
            xi = xi
        End If

    Next xi

Remove_Pre_LineFeeds:
    If DCount("RevId", "ATFRevT", "Comment Like '" & Chr(13) & Chr(10) & "*'") Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE ATFRevT SET ATFRevT.Comment = Replace([Comment],'" & Chr(13) & Chr(10) & "','',1,1) WHERE ATFRevT.Comment Like '" & Chr(13) & Chr(10) & "*'")
        DoCmd.SetWarnings True
    End If
    If DCount("RevId", "ATFRevT", "Comment Like '" & Chr(13) & Chr(10) & "*'") Then GoTo Remove_Pre_LineFeeds
    Exit Function
End Function
Public Function RemoveRichText(varText As Variant) As Variant
'Need refernence to Microsoft VBscript regular expressions
    Dim rgx    As New VBScript_RegExp_55.RegExp
    If Not IsNull(varText) Then
        rgx.Pattern = "<[^>]+>"
        'Global: Sets a Boolean value or returns a Boolean value that indicates
        'whether a pattern must match all the occurrences in a whole search string,
        'or whether a pattern must match just the first occurrence
        rgx.Global = True
        'Multiline: enables the regular expression engine to handle an input string that consists
        'of multiple lines. It changes the interpretation of the ^ and $ language elements so that
        'they match the beginning and end of a line, instead of the beginning and end of the input string.
        rgx.MultiLine = True
        varText = rgx.Replace(varText, "")
    End If
    varText = Replace(varText, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", " ")
    varText = Replace(varText, "&nbsp;", " ")
    RemoveRichText = varText
End Function
Sub DoSql(ReplaceX As String, ReplaceWith As String)
    Dim SQLz   As String
    SQLz = "UPDATE ATFRevT SET Comment = Replace([Comment],'" & ReplaceX & "','" & ReplaceWith & "',1,1) WHERE Comment Like '*" & ReplaceX & "*'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub
Sub DoSqlReason(ReplaceX As String, ReplaceWith As String)
    Dim SQLz   As String
    SQLz = "UPDATE ATF_ChangeReasons SET Reason = Replace(Reason,'" & ReplaceX & "','" & ReplaceWith & "',1,1) WHERE Reason Like '*" & ReplaceX & "*'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub

Public Function SetSessionProfile(GroupIDx As Long, formName As String)
    If DLookup("MLOrig", "vw_FiveKeys", "GroupID=" & GroupIDx) <> "FA" Then Exit Function
    SessionInt = 0
    SessionInt = Nz(DLookup("SessionID", "ATF_FAID_Sessions", "FormName ='" & formName & "' AND SysUser='" & FAUserName & "' AND isnull(Closed) AND GroupId=" & GroupIDx), 0)
    If SessionInt > 0 Then Exit Function
    Dim SQLz   As String
    SQLz = SQLz & "INSERT INTO ATF_FAID_Sessions (SchID, GroupID, sUnitCost, sChildCost, sGroupOEC, sSyndRent, sSchRent, sPdRent, sUTA, UnitsInGroup, FormName,  SysUser, Opened )"
    SQLz = SQLz & "SELECT OEC.SchID, OEC.GroupID, OEC.GroupOEC, OEC.GroupPartCost, nz(OEC.TotalGroupOEC,0), nz(OEC.SyndRent,0), nz(OEC.SchRent,0), nz(OEC.PerdiemRent,0), nz(UTASynd,0), OEC.Units "
    SQLz = SQLz & ",'" & formName & "' "
    SQLz = SQLz & ",'" & FAUserName & "' "
    SQLz = SQLz & ",'" & Now() & "' "
    SQLz = SQLz & "From vw_OEC_Rents_LRF_Cost As OEC WHERE OEC.GroupID=" & GroupIDx & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SessionInt = Nz(DMax("SessionID", "ATF_FAID_Sessions", "GroupID=" & GroupIDx), 0)
End Function

Public Function CompareSessionProfile(SessionX As Long, Item As String) As String
'(GroupIDx As Long, FormName, FaUser As String, Item As String)
    Dim SQLz   As String
    Dim TempColl As Collection
    Dim RST    As DAO.Recordset
    SQLz = SQLz & "SELECT W.SessionID, W.SysUser, W.FormName, W.Opened, N.SchID, N.GroupID "
    SQLz = SQLz & ", W.sUnitCost, N.GroupOEC AS UnitCost, IIf([sUnitCost]<>[GroupOEC],1,0) AS UnitCostChange "
    SQLz = SQLz & ", W.sSyndRent, N.SyndRent, IIf([sSyndRent]<>[SyndRent],1,0) AS SyndRentChange "
    SQLz = SQLz & ", W.sSchRent, N.SchRent, IIf([W.sSchRent]<>[N.SchRent],1,0) AS SchRentChange "
    SQLz = SQLz & ", W.sChildCost, N.GroupPartCost AS ChildCost, IIf([GroupPartCost]<>[sChildCost],1,0) AS ChildCostChange "
    SQLz = SQLz & ", W.sPdRent, N.PerdiemRent AS PdRent, IIf([sPdRent]<>[PerdiemRent],1,0) AS PDRentChange "
    SQLz = SQLz & ", W.UnitsInGroup AS sUnitsInGroup, N.Units AS UnitsInGroup, IIf([UnitsInGroup]<>[Units],1,0) AS UnitsInGroupChange "
    SQLz = SQLz & "FROM ATF_FAID_Sessions AS W INNER JOIN vw_OEC_Rents_LRF_Cost AS N ON W.GroupID = N.GroupID "
    SQLz = SQLz & "WHERE W.SessionID= " & SessionX & ";"
    'SQlz = SQlz & "WHERE W.SysUser='" & FAUserName & "' AND W.FormName='" & FormName & "' AND N.GroupID=" & GroupIDx & ";"
    On Error GoTo FixMe
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount = 0 Then GoTo FixMe

    Set TempColl = New Collection
    TempColl.Add Key:="WasUnitCost", Item:=RST!sUnitCost
    TempColl.Add Key:="IsUnitCost", Item:=RST!Unitcost
    TempColl.Add Key:="CostChanged", Item:=RST!UnitCostChange

    TempColl.Add Key:="WasChildCost", Item:=RST!sChildCost
    TempColl.Add Key:="IsChildCost", Item:=RST!ChildCost
    TempColl.Add Key:="ChildCostChanged", Item:=RST!ChildCostChange

    TempColl.Add Key:="WasSyndRent", Item:=RST!sSyndRent
    TempColl.Add Key:="IsSyndRent", Item:=RST!SyndRent
    TempColl.Add Key:="SyndRentChanged", Item:=RST!SyndRentChange

    TempColl.Add Key:="WasSchRent", Item:=RST!sSchRent
    TempColl.Add Key:="IsSchRent", Item:=RST!SchRent
    TempColl.Add Key:="SchRentChanged", Item:=RST!SchRentChange

    TempColl.Add Key:="WasPdRent", Item:=RST!sPdRent
    TempColl.Add Key:="IsPdRent", Item:=RST!PdRent
    TempColl.Add Key:="PdRentChanged", Item:=RST!PdRentChange

    TempColl.Add Key:="WasUnitsInGroup", Item:=RST!sUnitsInGroup
    TempColl.Add Key:="IsUnitsInGroup", Item:=RST!UnitsInGroup
    TempColl.Add Key:="UnitsInGroupChange", Item:=RST!UnitsInGroupChange

    CompareSessionProfile = TempColl(Item)
    ' Debug.Print TempColl("WasUnitCost")
    RST.close
    Set RST = Nothing
    Set TempColl = Nothing
    Exit Function
FixMe:
    RST.close
    Set RST = Nothing
    Set TempColl = Nothing
End Function
Public Function PrintEmailPO_Vendor(PONumberX As String, GroupIDx As Long)
    Dim FileToAttachment As String
    Dim FontSetting, msg1 As String
    Dim SchIDx As Integer
    SchIDx = DLookup("SchId", "vw_sixKeys", "SGrpId=" & GroupIDx)
    FileToAttachment = "C:\Faas\" & FAUserName & "\Temp\PO_" & DLookup("ClientShortName", "vw_sixKeys", "Schid=" & SchIDx) & "_Sch_" & DLookup("Schedule", "vw_sixKeys", "Schid=" & SchIDx) & "-" & Format(Date, "mm-dd-yy") & ".pdf"
    DoCmd.OutputTo acOutputReport, "PO_VendorBoth", acFormatPDF, FileToAttachment, False

SendNotice:
    'make email message
    FontSetting = "<span style='font-size:10pt;font-family:" & Chr(34) & "Segoe UI" & Chr(34) & ",sans-serif'>"
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'><b>ATF Imported Notice</b></p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:13px'>To: ATF Distribution<br /><br />"
    msg1 = msg1 & "Re. Client:  <B>" & DLookup("ClientShortName", "vw_SixKeys", "SchID=" & SchIDx) & "</B><br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;MLNo: <B>" & DLookup("MLNo", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Schedule Number: <B>" & DLookup("Schedule", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Number of Groups: <B>" & DCount("SGrpSchID", "SchGrp", "SGrpSchId=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Total Units in ATF: <B>" & DCount("UnitID", "vw_SixKeys", "SchID=" & SchIDx) & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Date of Import: <B> " & DLookup("ChangeDate", "ATF_ChangeReasons", "Action ='Schedule Import' AND Schid=" & SchIDx & "") & "</B><Br /> "
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Import ran by: <B> " & FAUserName & "</B><Br />"
    msg1 = msg1 & "&emsp;&nbsp;&nbsp;Please see the attached PDF"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Dim EmailToList As String
    'EmailToList = "kzeigler@FleetAdvantage.com;pachille@FleetAdvantage.com;bmcmahon@FleetAdvantage.com;JGriffin@FleetAdvantage.com;rwasserman@FleetAdvantage.com"
    ' EmailToList = "ATFDistribution@FleetAdvantage.com"
    Call SendEmailMessage("", "ATLAAS Core Message: ATF Import", msg1, False, True, False, "ljohnson@fleetadvantage.com", , , FileToAttachment)

End Function
