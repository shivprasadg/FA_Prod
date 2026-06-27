Attribute VB_Name = "ATLAASMessager"
Option Compare Database
Option Explicit
    Public msg1 As String
    Public EmailToGrp As String
    Public EmailCCGrp As String

Public Function AllUnitsInGroup_InvoicedMessage(GroupIDx As Long)
    Dim ClientX, Schx, Grpx, AcceptDateX, InvoiceDateX, SQLy, BLCDBasis, PerDiemBasis, MSOx, subjectx As String
    Dim ClientIDint, SchIDInt As Integer
     InvoiceDateX = DMax("UnitInvoiceDate", "Units", "GroupRef=" & GroupIDx)
     If CDate(InvoiceDateX) < CDate("10/1/2017") Then Exit Function
'***** check if notice was already sent
    If IsDate(DLookup("NoticeActionDate", "ATLAASNotifications", "GroupID=" & GroupIDx & " AND NoticeTypeID=4")) Then
        DoCmd.SetWarnings True
        Exit Function
    End If
    Schx = DLookup("Schedule", "vw_SixKeys", "SGrpID=" & GroupIDx)
    Grpx = DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & GroupIDx)
    AcceptDateX = DMax("UnitAcceptDate", "Units", "GroupRef=" & GroupIDx)
    ClientIDint = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    SchIDInt = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    If AllUnitsInGroup_MsoResv(GroupIDx) = True Then MSOx = "Y" Else MSOx = "N"
    BLCDBasis = DLookup("StartsOn", "Lookup_LeaseStartsOn", "ID=" & Nz(DLookup("SgrpCommInd", "Schgrp", "SGrpID=" & GroupIDx), 10))
    PerDiemBasis = DLookup("StartsOn", "Lookup_LeaseStartsOn", "ID=" & Nz(DLookup("PerDiemStartsOn", "Schgrp", "SGrpID=" & GroupIDx), 10))
    SQLy = "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) " & _
           "SELECT " & ClientIDint & ", " & SchIDInt & ", " & GroupIDx & ", #" & Date & "#, #" & Nz(DMax("UnitInvoiceDate", "Units", "GroupRef=" & GroupIDx), Date) & "#, " & FAUserIDx & ", 'AllInvoiced',4,1;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True

    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>All Units Invoiced</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Syndication Dept.<br /><br />"
    msg1 = msg1 & "Client: <B>" & ClientX & "</B><br />"
    msg1 = msg1 & "Master Lease: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Schedule: <B>" & Schx & "</B><br />"
    msg1 = msg1 & "Group: <B>" & Grpx & "</B><br />"
    msg1 = msg1 & "BLCD Basis:<B> " & BLCDBasis & "</B><br />"
    msg1 = msg1 & "BLCD:<B> " & DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Per Diem Basis:<B> " & PerDiemBasis & "</B><br />"
    msg1 = msg1 & "All MSOs Received:<B> " & MSOx & "</B><br />"
    msg1 = msg1 & "Units In Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "All Units Invoiced As Of: <B>" & InvoiceDateX & "</B><br />"
    msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"

    subjectx = "ATLAAS Core Message: All Units Invoiced Client: " & ClientX & " Sch: " & Schx & " Grp: " & Grpx & ""
'submit emails  as if it/they are an array
    EmailToGrp = "@(""<kzeigler@fleetadvantage.com>"",""<kbabb@FleetAdvantage.com>"",""<lredmond@FleetAdvantage.com>"",""<shora@FleetAdvantage.com>"")"
    EmailCCGrp = "@(""<jWharton@fleetadvantage.com>"",""<ssweet@fleetadvantage.com>"",""<ljohnson@fleetadvantage.com>"")"
    Call SendPowerShellEmail(EmailToGrp, subjectx, msg1, EmailCCGrp)

End Function

Public Function AllUnitsInGroup_AcceptedMessage(GroupIDx As Long)
    Dim ClientX, Schx, Grpx, AcceptDateX, SQLy, BLCDBasis As String
    Dim ClientIDint, SchIDInt As Integer

    ' check if notice was already sent
    If IsDate(DLookup("NoticeActionDate", "ATLAASNotifications", "GroupID=" & GroupIDx & " AND NoticeTypeID=1")) Then
        DoCmd.SetWarnings True
        Exit Function
    End If


    Schx = DLookup("Schedule", "vw_SixKeys", "SGrpID=" & GroupIDx)
    Grpx = DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & GroupIDx)
    AcceptDateX = DMax("UnitAcceptDate", "Units", "GroupRef=" & GroupIDx)
    ClientIDint = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    SchIDInt = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    BLCDBasis = DLookup("StartsOn", "Lookup_LeaseStartsOn", "ID=" & DLookup("SgrpCommInd", "Schgrp", "SGrpID=" & GroupIDx))

    SQLy = "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) " & _
           "SELECT " & ClientIDint & ", " & SchIDInt & ", " & GroupIDx & ", #" & Date & "#, #" & Nz(DMax("UnitAcceptDate", "Units", "GroupRef=" & GroupIDx), Date) & "#, " & FAUserIDx & ", 'AllAccepted',1,1;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True

    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>All Units Accepted</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Docs Dept.<br /><br />"
    'Msg1 = Msg1 & "All of the units in Group (" & GrpX & ") have been marked Status 'A':<br /><br />"
    msg1 = msg1 & "Client: <B>" & ClientX & "</B><br />"
    msg1 = msg1 & "Master Lease: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Schedule: <B>" & Schx & "</B><br />"
    msg1 = msg1 & "Group: <B>" & Grpx & "</B><br />"
    msg1 = msg1 & "BLCD Basis:<B> " & BLCDBasis & "</B><br />"
    msg1 = msg1 & "BLCD:<B> " & DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Units In Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "All Units Accepted As Of: <B>" & AcceptDateX & "</B><br />"
    'Msg1 = Msg1 & "Last In Service Date: <B>" & DMax("UnitInServ", "Units", "GroupRef=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"

    EmailToGrp = "@(""<pachille@FleetAdvantage.com>"")"
    EmailCCGrp = "@(""<ljohnson@fleetadvantage.com>"",""<ssweet@fleetadvantage.com>"")"
    Call SendPowerShellEmail(EmailToGrp, "ATLAAS Core Message: All Units Accepted", msg1, EmailCCGrp)
    'Call SendEmailMessage("pachille@FleetAdvantage.com", "ATLAAS Core Message: All Units Accepted", msg1, True, True, False, "JGriffin@FleetAdvantage.com;ljohnson@fleetadvantage.com")

End Function
Public Function AllUnitsInGroup_PerDiemSet_Message(GroupIDx As Long, Optional ReportDate As String)

    If IsDate(DLookup("NoticeActionDate", "ATLAASNotifications", "GroupID=" & GroupIDx & " AND NoticeTypeID=2")) Then Exit Function

    Dim ClientX, Schx, Grpx, AcceptDateX, SQLy, AllAccepted, AllPDSet, PerDiemBasis As String
    Dim ClientIDint, SchIDInt, x As Integer

    If AllUnitsInGroup_PerDiemSet(GroupIDx) = True Then AllPDSet = "Y" Else AllPDSet = "N"
    DoCmd.SetWarnings True


    Schx = DLookup("Schedule", "vw_SixKeys", "SGrpID=" & GroupIDx)
    Grpx = DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & GroupIDx)
    AcceptDateX = DMax("UnitAcceptDate", "Units", "GroupRef=" & GroupIDx)
    ClientIDint = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    SchIDInt = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    PerDiemBasis = DLookup("StartsOn", "Lookup_LeaseStartsOn", "ID=" & DLookup("PerDiemStartsOn", "Schgrp", "SGrpID=" & GroupIDx))

    SQLy = "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) " & _
           "SELECT " & ClientIDint & ", " & SchIDInt & ", " & GroupIDx & ", #" & Date & "#, #" & DMax("UnitPDStartDate", "Units", "GroupRef=" & GroupIDx) & "#, " & FAUserIDx & ", 'All Perdiem Set',2,1;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True

    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'><b>All Units Per Diem Set</b></p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Docs Dept.<br /><br />"
    'Msg1 = Msg1 & "All of the units in Group (" & GrpX & ") have been marked Status 'A':<br /><br />"
    msg1 = msg1 & "Client: <B>" & ClientX & "</B><br />"
    msg1 = msg1 & "Master Lease: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Schedule: <B>" & Schx & "</B><br />"
    msg1 = msg1 & "Group: <B>" & Grpx & "</B><br />"
    msg1 = msg1 & "Per Diem Basis:<B> " & PerDiemBasis & "</B><br />"
    'Msg1 = Msg1 & "BLCD:<B> " & DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Units In Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "All Units Accepted As Of: <B>" & AcceptDateX & "</B><br />"
    msg1 = msg1 & "All Units Per Diems Set On: <B>" & Date & "</B><br />"
    'Msg1 = Msg1 & "Last In Service Date: <B>" & DMax("UnitInServ", "Units", "GroupRef=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"

    EmailToGrp = "@(""<lredmond@FleetAdvantage.com>"",""<kzeigler@FleetAdvantage.com>"")"
    EmailCCGrp = "@(""<ljohnson@fleetadvantage.com>"")"
     Call SendPowerShellEmail(EmailToGrp, "ATLAAS Core Message: All Units Per Diem Set", msg1, EmailCCGrp)
    'Call SendEmailMessage("lredmond@FleetAdvantage.com;kzeigler@FleetAdvantage.com", "ATLAAS Core Message: All Units Per Diem Set", msg1, True, True, False, "JGriffin@FleetAdvantage.com;ljohnson@fleetadvantage.com")
    Exit Function
FixMe:
    RST.close
    Set RST = Nothing
End Function
Public Function Remote_Call_BLCD()
    Call BLCDSyndFinalMessage(0)
End Function
Public Function OpenGroupID()
Dim FileNum As Integer
Dim DataLine As String
DoCmd.close acForm, "GroupForm"
FileNum = FreeFile()
Open "c:\FAAS\GroupID.txt" For Input As #FileNum
While Not EOF(FileNum)
    Line Input #FileNum, DataLine ' read in data 1 line at a time
    DoCmd.OpenForm ("GroupForm"), , , "sGrpID=" & Int(DataLine)
Wend

End Function
Public Function BLCDSyndFinalMessage(GroupIDx As Long, Optional ReportDate As String)
    Dim ClientX, Schx, AllAccepted, AllPDSet, SQLy As String
    Dim ClientIDint, SchIDInt, x As Integer
    DoCmd.SetWarnings True
    If GroupIDx = 0 Then GoTo SendNotice
    If IsDate(DLookup("NoticeDate", "ATLAASNotifications", "GroupID=" & GroupIDx & " AND NoticeTypeID=3")) Then GoTo SendNotice
    Schx = DLookup("Schedule", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientIDint = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    SchIDInt = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)

    SQLy = "INSERT INTO ATLAASNotifications ( clientID, SchID, GroupID, NoticeDate, NoticeActionDate, UserID, NoticeType, NoticeTypeID, NoticeSent ) " & _
           "SELECT " & ClientIDint & ", " & SchIDInt & ", " & GroupIDx & ", #" & Date & "#, " & _
           "#" & DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx) & "#, " & FAUserIDx & ", 'BLCD Final',3,0;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLy)
    DoCmd.SetWarnings True

SendNotice:
    If GroupIDx <> 0 Then Exit Function
    If ReportDate & "test" = "test" Then ReportDate = Date
    Dim RST    As DAO.Recordset
    SQLy = "Select * From ATLAASNotifications WHERE NoticeSent=0 AND NoticeTypeID = 3 AND NoticeDate <= #" & ReportDate & "# Order By GroupID;"
    Set RST = CurrentDb.OpenRecordset(SQLy, dbOpenSnapshot)
    On Error GoTo FixMe
    RST.MoveFirst
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0

    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>BLCD Set</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Accounting - Billing Dept.<br /><br />"
    msg1 = msg1 & "Syndication was finalized and the BLCD was entered for:<br /><br />"

    If DLookup("BLCD", "Schgrp", "sGrpId=" & GroupIDx & "") < DLookup("AssignmentDate", "Schgrp", "sGrpId=" & GroupIDx & "") Then
     msg1 = msg1 & "<p style='color:Red;font-family:Segoe UI;font-size:20px;'>This Group has STRIP RENTS!</p><br />"
    End If

    msg1 = msg1 & "<table style='font-family:Segoe UI;font-size:12px' BORDER='1'    WIDTH='50%'   CELLPADDING='2' CELLSPACING='1' border-collapse='1' >"
    msg1 = msg1 & "<TR ALIGN='Center'><TH>Client</TH><TH>Lease No</TH><TH>Schedule</TH><TH>Group</TH><TH>BLCD</TH><TH>Units</TH><TH>All Accepted</TH><TH>All PD Set</TH><TH>FA User</TH><TH>Date Set</TH></TR>"
    For x = 1 To RST.RecordCount
        If AllUnitsInGroup_Accepted(RST!groupID) = True Then AllAccepted = "Y" Else AllAccepted = "N"
        If AllUnitsInGroup_PerDiemSet(RST!groupID) = True Then AllPDSet = "Y" Else AllPDSet = "N"

        msg1 = msg1 & "<TD>" & DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD>" & DLookup("Schedule", "vw_SixKeys", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD>" & DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD ALIGN='Center'>" & DLookup("BLCD", "SchGrp", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD ALIGN='Center'>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & RST!groupID) & "</TD> " & _
               "<TD ALIGN='Center'>" & AllAccepted & "</TD>" & _
               "<TD ALIGN='Center'>" & AllPDSet & "</TD>" & _
               "<TD ALIGN='Left'>" & Nz(DLookup("[Employee Name]", "TBL_Employees", "ID=" & RST!USERID), "Atlaas Core") & "</TD>" & _
               "<TD ALIGN='Center'>" & RST!NoticeDate & "</TD></TR>"
        SQLy = "UPDATE ATLAASNotifications SET NoticeSent = 1 WHERE NoticeTypeID = 3 AND GroupID=" & RST!groupID & ";"
        DoCmd.SetWarnings False
     '   DoCmd.RunSQL (SQLy)
        DoCmd.SetWarnings True
        RST.MoveNext
    Next x
    RST.close
    Set RST = Nothing

    EmailToGrp = "@(""<hprikes@FleetAdvantage.com>"",""<sroyer@FleetAdvantage.com>"",""<lredmond@FleetAdvantage.com>"")"
    EmailCCGrp = "@(""<pachille@FleetAdvantage.com>"",""<ljohnson@fleetadvantage.com>"")"
    Call SendPowerShellEmail(EmailToGrp, "ATLAAS Core Message: BLCD Set", msg1, EmailCCGrp)
   ' Call SendEmailMessage("hprikes@FleetAdvantage.com;sroyer@FleetAdvantage.com;lredmond@FleetAdvantage.com", "ATLAAS Core Message: BLCD Set", msg1, True, True, False, "pachille@FleetAdvantage.com;jgriffin@fleetadvantage.com;ljohnson@fleetadvantage.com")
    Exit Function
FixMe:
  msg1 = msg1 & "</table>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"

       RST.close
    Set RST = Nothing
End Function
Public Function MessageUser(XCommand As String, Optional msg1 As String, Optional Msg2 As String, Optional Rpaint As Boolean)
    If XCommand = "Open" Then
        DoCmd.close acForm, "MessagePopUp"
        DoCmd.OpenForm "MessagePopUp", acNormal
    End If
    If XCommand = "Close" Then
        DoCmd.close acForm, "MessagePopUp"
        Exit Function
    End If
    Forms!MessagePopUp.Message1.caption = Nz(msg1, Forms!MessagePopUp.Message1.caption)
    Forms!MessagePopUp.Message2.caption = Nz(Msg2, Forms!MessagePopUp.Message2.caption)
    Forms!MessagePopUp.Repaint
End Function
Public Function AllUnitsInGroup_PerDiemSet(GroupIDx As Long) As Boolean
    AllUnitsInGroup_PerDiemSet = False
    If DMin("Nz([UnitPDStartDate],#1/1/1900#)", "Units", "GroupRef=" & GroupIDx) > "1/1/1900" Then AllUnitsInGroup_PerDiemSet = True
End Function
Public Function AllUnitsInGroup_Accepted(GroupIDx As Long) As Boolean
    AllUnitsInGroup_Accepted = False
    If DMax("Status", "vw_SixKeys", "SGrpID=" & GroupIDx) = "A" Then AllUnitsInGroup_Accepted = True
End Function
Public Function AllUnitsInGroup_Invoiced(GroupIDx As Long) As Boolean
    AllUnitsInGroup_Invoiced = False
    If Nz(DMin("UnitInvoiceDate", "Units", "GroupRef=" & GroupIDx), "1/1/1900") <> "1/1/1900" Then AllUnitsInGroup_Invoiced = True
End Function
Public Function AllUnitsInGroup_MsoResv(GroupIDx As Long) As Boolean
    AllUnitsInGroup_MsoResv = False
        If Nz(DMin("UnitMSOCopy", "Units", "GroupRef=" & GroupIDx), Nz(DMin("UnitMSORecd", "Units", "GroupRef=" & GroupIDx), "1/1/1900")) <> "1/1/1900" Then AllUnitsInGroup_MsoResv = True
End Function
Public Function IsGroupBLCDSet(GroupIDx As Long) As Boolean
    IsGroupBLCDSet = False
    If DLookup("Assigned_Indicator", "SchGrp", "SGrpID=" & GroupIDx) = "F" _
       And IsDate(DLookup("BLCD", "vw_SixKeys", "MlOrig='FA' AND SGrpID=" & GroupIDx)) Then
        IsGroupBLCDSet = True
    End If
End Function

Public Function WasGroupBLCDNoticeSent(GroupIDx As Long) As Boolean
    WasGroupBLCDNoticeSent = False
    If IsDate(DLookup("NoticeActionDate", "ATLAASNotifications", "GroupID=" & GroupIDx & " AND NoticeTypeID=3")) Then
        WasGroupBLCDNoticeSent = True
    End If
End Function
Public Function CheckIfSendBLCDEmailSent() As Boolean
    CheckIfSendBLCDEmailSent = True
    If DCount("GroupID", "ATLAASNotifications", "NoticeSent=0 AND NoticeTypeID=3") > 0 Then
        CheckIfSendBLCDEmailSent = False
    End If
End Function

Public Function SQLMail()
Dim task As String
Dim msg1 As String
Dim qdef As DAO.QueryDef

task = "Writing FRU-Pending to Staging"
msg1 = "<p style=''color:Navy;font-family:Segoe UI;font-size:20px;''>Scheduled Business ETL Failed</p></p>"
msg1 = msg1 & "<p style=''font-family:Segoe UI;font-size:13px''>Task that was running:" & task & " which failed at:<b> " & Now() & " </b></p>"

Set qdef = CurrentDb.CreateQueryDef("")
'qdef.Connect = "ODBC;DRIVER=SQL Server;SERVER=FA-Phoenix;DATABASE=FleetAdvantage_beSQL;Trusted_Connection=yes"
qdef.Connect = CurrentDb.TableDefs("units").Connect
qdef.SQL = "msdb.DbO.sp_send_dbmail @profile_name = 'PhoenixEmail'"
qdef.SQL = qdef.SQL & ",@recipients = 'ljohnson@fleetadvantage.com'"
'qdef.SQL = qdef.SQL & ",@copy_recipients = 'it@fleetadvantage.com'"
qdef.SQL = qdef.SQL & ",@subject = 'Scheduled business ETL worked'"
qdef.SQL = qdef.SQL & ",@body = " & msg1
qdef.SQL = qdef.SQL & ",@body_format = 'HTML'"
qdef.ReturnsRecords = False
qdef.Execute
End Function
