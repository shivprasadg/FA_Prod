Attribute VB_Name = "Trust_Synd_Email_Funtions"
Option Compare Database
Dim SQLz       As String

Function MakeCustodianFiles(iSchID As Integer)
    Dim Countx As Integer
    Dim DateX, TrackX As String
    Dim Response As String
    Dim CloseMe As String
    Dim qdf As DAO.QueryDef
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "TList"
    DoCmd.DeleteObject acQuery, "DataFile"
    On Error GoTo 0

    SQLz = "SELECT CTT.[FA TitleLog Num] AS [FA-Asset Num]"
    SQLz = SQLz & ", CTT.VIN AS Expr1, CTT.[File Type] AS Expr2"
    SQLz = SQLz & ", CTT.[Portfolio Certificate No] AS Expr3"
    SQLz = SQLz & ", CTT.[Lender or Holder Name] AS Expr4 "
    SQLz = SQLz & "FROM TList As CTT Order by [FA TitleLog Num];"

    Set qdf = CurrentDb.CreateQueryDef("TList", "Select * From vw_CustodianTitleTrust2 Where uStatus <>'C' and SchID=" & iSchID & " Order By [Leasee Schedule No], [FA TitleLog Num] ;")
    Set qdf = CurrentDb.CreateQueryDef("DataFile", SQLz)
    Countx = DCount("Lessee", "TList")
    'GoTo AskDate
    If Countx < 1 Then
        MsgBox "There is missing data this is preventing this report from running," & vbNewLine & "(Such as a missing Title Locations, Bank or Portfolio #).", vbCritical, "No records"
        Exit Function
    End If
    DateX = Nz(DLookup("DateSent", "TList"), Date)
    Response = MsgBox("Do you want to send an Custodian Submission Email?", vbQuestion + vbYesNo + vbCritical, "Email it?")
    If Response = vbYes Then CloseMe = "Close" Else CloseMe = "Open"
    ExcelCustodian2 "DataFile", "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2DataFile.xlsx", "DataFile", CloseMe, , Countx, "Data Files - Sent " & DateX, False, "CC"
    ExcelCustodian1 "TList", "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2.xlsx", "Custodian", CloseMe, , Countx, "Equipment Files - Sent " & DateX, False, "CC"
    If Response = vbYes Then Call EmailToTrust(iSchID)
AskDate:
    Response = MsgBox("Do you want to log these " & Countx & " Titles as 'Sent To Custodian'?", vbQuestion + vbYesNo + vbCritical, "Log it?")
    If Response = vbYes Then
        Dim DateSent As String

        DateSent = InputBox("Provide the 'Sent Date' for these titles." & vbNewLine & _
            "          Please enter that date, below." & vbNewLine & _
            "          Today's date is the default." & vbNewLine & vbNewLine & _
            "          ( You must use a date )" _
            , "Custodian Titles, Sent Date", Date)
        If Not IsDate(DateSent) Then GoTo AskDate
        SQLz = "INSERT INTO UnitTitleLog ( UnitRef, UnitTitleLogDate, EntryTypeId, EntryReasonId, TitleSt, TitleNumber, EntryReason, UserID )" & _
            "SELECT Units.UnitID, #" & [DateSent] & "# AS SentDate, 2 AS Type, 4 AS Reasonid, CT.[State of Title], CT.TitleNum, 'Sent To Custodian' AS Reason " & _
            ",'" & FaUserName & "' " & _
            "FROM vw_CustodianTitleTrust2 As CT INNER JOIN Units ON CT.VIN = Units.UnitVIN " & _
            "WHERE SchID=" & iSchID & " ORDER BY Units.UnitID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        MsgBox "Title Log Updated For " & Countx & " Units", vbInformation, "Log Updated"
    End If
    DoCmd.DeleteObject acQuery, "TList"
    DoCmd.DeleteObject acQuery, "DataFile"
End Function
Function EmailToTrust(iSchID As Integer)
    Dim x, Y, iBankId, cUnits As Integer
    x = DCount("BankId", "vw_CustodianTitleTrust2", "SchID=" & iSchID)
    If x < 1 Then
        MsgBox "There are no Units to send", , "No Stuff"
        Exit Function
    End If

    Dim MasterLeaseX, Contactx, ContactEmailX, CcContactX, CcEmailX, LeaseeX, LenderX, ScheduleX, _
        PortfolioX, subjectx, SubjectX2, UserX, UserEmailX, msg1, ItemsX, Linkx(2) As String
    Dim RST    As DAO.Recordset
    ' Get causality units if any
    cUnits = DCount("uStatus", "vw_CustodianTitleTrust2", "uStatus ='C' and SchID=" & iSchID & "")

    '*************************************** Get Schedule Details
    SQLz = "SELECT * From vw_CustodianTitleTrust2 Where SchID=" & iSchID
    SQLz = SQLz & " Order By [Leasee Schedule No], [FA TitleLog Num] "
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveFirst
    iBankId = RST!BankID
    LenderX = RST![Lender or Holder Name]
    LeaseeX = RST!Lessee
    MasterLeaseX = RST![Leasee Master Lease Number]
    ScheduleX = RST![Leasee Schedule No]
    PortfolioX = RST![Portfolio Certificate No]
    RST.close
    '*************************************** Get the CustodianEmails
    Contactx = ""
    ContactEmailX = ""
    SQLz = "SELECT * From BankCustEmails"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Contactx = ""
    ContactEmailX = ""
    If RST.RecordCount > 0 Then
        RST.MoveFirst
        Do Until RST.EOF
            Contactx = RST!CustName
            ContactEmailX = RST!CustEmail & ";"
            If RST!CustCName & "test" <> "test" Then
                Contactx = Contactx & " - " & RST!CustCName
                ContactEmailX = ContactEmailX & RST!CustCEmail & ";"
            End If
            Contactx = Contactx & " - " & RST!ABSEmail
            ContactEmailX = ContactEmailX & RST!ABSEmail
            RST.MoveNext
        Loop
    End If

    '*************************************** Get The Incumbent Emails
    Dim PTIName, PTIEmail As String
    Dim TICName, TICEmail As String
    Dim TCName, TCEmail As String
    PTIName = ""
    PTIEmail = ""
    SQLz = "SELECT LC.clientRef"
    SQLz = SQLz & ", LC.[contactFullName] AS TrustContactName"
    SQLz = SQLz & ", LC.[contactEmail] AS TrustContactEmail"
    SQLz = SQLz & ", LC.contactClassification "
    SQLz = SQLz & "FROM Location_contacts AS LC "
    SQLz = SQLz & "WHERE LC.contactEmail Is Not Null "
    SQLz = SQLz & "AND (LC.contactClassification Like 'TR*' OR LC.contactClassification Like '*TRUST*') "
    SQLz = SQLz & "AND clientRef=" & iBankId
    SQLz = SQLz & " ORDER BY [contactFullName]"

    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveFirst
        Do Until RST.EOF
            If RST!contactClassification = "Primary Trust Incumbent" Then
                PTIName = PTIName & RST!TrustContactName & " - "
                PTIEmail = PTIEmail & RST!TrustContactEmail & ";"
            End If
            If RST!contactClassification Like "TR Incumbent*" Then
                TICName = TICName & RST!TrustContactName & " - "
                TICEmail = TICEmail & RST!TrustContactEmail & ";"
            End If
            If RST!contactClassification = "TR Contact" Then
                TCName = TCName & RST!TrustContactName & " - "
                TCEmail = TCEmail & RST!TrustContactEmail & ";"
            End If
            RST.MoveNext
        Loop
    End If
    RST.close
    Set RST = Nothing
    '*************************************** Set Email Order
    CcContactX = PTIName & TICName & TCName
    CcContactX = Left(CcContactX, (Len(CcContactX) - 3))
    CcEmailX = PTIEmail & TICEmail & TCEmail
    CcEmailX = Left(CcEmailX, (Len(CcEmailX) - 1))

    '*************************************** Build the Email
    subjectx = "FA Title Trust - Equipment Files:(" & PortfolioX & ") Sending out (" & (x - cUnits) & ") Titles " & Date & " Lease No. " & MasterLeaseX & "; Schedule No. " & ScheduleX & ""
    SubjectX2 = "Subject: FA Title Trust - Equipment Files: (<B>" & PortfolioX & "</B>) Sending out (<B>" & (x - cUnits) & "</B>) Titles " & Date & "</B><br />Re: Lease No. <B>" & MasterLeaseX & "</B>; Schedule No.: <B>" & ScheduleX & "</B>"
    UserX = Forms!FAMainMenu![USERID]
    UserEmailX = Forms!FAMainMenu![USERID] & "@fleetadvantage.com"
    'CC-ed
    'Primary Trust Incumbent
    'Trust Incumbent/Correspondent
    'Trust Contact
    If PTIName & "test" = "test" Then PTIName = "<B>[<i>Insert Target Recipient Here</i>]</b>"
    If Contactx Like "*Sandra Hora*" Then Contactx = Replace(Contactx, "Sandra Hora -", "")
    msg1 = ""
    Linkx(0) = ""
    Linkx(1) = ""
    Linkx(0) = "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2.xlsx"
    Linkx(1) = "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2DataFile.xlsx"
    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Calibri;font-size:20pt;'><b>Custodian Title Trust Submission</b></p>"
    msg1 = msg1 & "<p style='font-family:Calibri;font-size:11pt'>"
    msg1 = msg1 & "To: <b>" & Contactx & "</b><br />"
    msg1 = msg1 & "CC<b> " & CcContactX & "</b><br />"
    msg1 = msg1 & SubjectX2 & "<br/><br/>"
    msg1 = msg1 & "Dear Custodian,<br/><br/>"
    msg1 = msg1 & "Please be advised that I am sending (<B>" & (x - cUnits) & "</B>) Equipment Files to your attention via Federal Express Tracking #: <B>[ <i>Input Tracking Number HERE</i> ]</B>.<br />"
    msg1 = msg1 & "with respect to Lease No. <B>" & MasterLeaseX & "</B>; Schedule No. <B>" & ScheduleX & "</B>  allocated to <B> " & LenderX & "</B> - Portfolio Interest Acct No. <B>" & PortfolioX & ".</B><br /><br />"
    msg1 = msg1 & "I have attached a <i><u>Master Spreadsheet</u></i> of the Equipment Files being sent (for convenience of reference purposes for all parties copied), "
    msg1 = msg1 & "along with the <i><u>Data File</u></i> with the pertinent information you require.<br /><br />"
    If cUnits > 0 Then
       msg1 = msg1 & "<b>Please Note:</b><Br />This Schedule originally had " & x & " units but there "
       If cUnits = 1 Then
            msg1 = msg1 & "is a causality unit in this Schedule<br />which is not included in this submission"
            Else
            msg1 = msg1 & "are " & cUnits & " causality units in this Schedule<br />which are not included in this submission"
            End If
       msg1 = msg1 & " which reduces the expected count of units to<b> " & (x - cUnits) & " units</b><Br /><Br />"
    End If
    msg1 = msg1 & "Once received and reconciled, kindly confirm your receipt and forward an executed Certificate to <B>" & PTIName & "</B> and kindly copy ALL recipients included in this email.<br/>"
    msg1 = msg1 & "<p style='font-family:Calibri; font-size:11pt; margin-left:0px'>"
    msg1 = msg1 & "Thank you and of course, if you have any questions please contact me.<br /><br />"
    msg1 = msg1 & "Very truly yours,<br />"
    'msg1 = msg1 & "email created: " & Date & " (via Fleet Advantage's ATLAAS System).<br/><br />"
    msg1 = msg1 & "</body></html>"

    Call SendEmailMessage(ContactEmailX, subjectx, msg1, False, True, False, CcEmailX, "", "", Linkx())

End Function
Public Function MarkGroupDraftAsPaid(GroupIDx As Long)

    Dim gFundDate As Date

    gFundDate = DLookup("FundDate", "SchGrp", "sGrpID=" & GroupIDx)

    SQLz = "UPDATE Units " & _
        "   SET UnitVendorPytDate = '" & gFundDate & "' " & _
        "   WHERE GroupRef = " & GroupIDx & " AND Nz(UnitIFAmt, 0) = 0"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)

    SQLz = "UPDATE Units " & _
        "   INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
        "   SET ChildParts.ChildVendorPytDate = '" & gFundDate & "' " & _
        "   WHERE Nz([cIFAmount],0)=0 AND Units.GroupREF=" & GroupIDx & ";"
    DoCmd.RunSQL (SQLz)

    SQLz = " UPDATE DrawDowns Set DrawPaidOff = '" & gFundDate & "'" & _
           " WHERE DrawDownID in (SELECT DrawDownID FROM Units " & _
           "                     WHERE Not DrawDownId Is Null " & _
           "                     AND GroupRef =" & GroupIDx & " Group By DrawDownID) "
    DoCmd.RunSQL (SQLz)

    SQLz = "UPDATE DrawDowns Set DrawPaidOff = '" & gFundDate & "' " & _
        "   WHERE DrawPaidOff Is Null " & _
        "   AND DrawDownID in (SELECT CP.DrawDownID " & _
        "                   FROM ChildParts as CP " & _
        "                   INNER JOIN Units AS U ON U.UnitID = CP.UnitREF" & _
        "                   WHERE Not CP.DrawDownId Is Null And u.GroupRef = " & GroupIDx & " " & _
        "                   GROUP By CP.DrawDownID)"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = ""
End Function
Public Function GroupFinalizedCheck(GroupIDx As Long, Optional NoDialog As Boolean, Optional NoEmail As Boolean) As String
    Dim RST    As DAO.Recordset
    Dim GFT    As DAO.Recordset

    Set RST = CurrentDb.OpenRecordset("Select * From SchGrp Where sGrpID=" & GroupIDx & "", dbOpenSnapshot)
    GroupFinalizedCheck = "Good"
    Dim AcctX, DocsX, OpsX, SyndX, TMTx, Msgx, Prefix, Suffix, SQLz As String
    Dim DearAcct, DearDocs, DearOps, DearSynd, DearTMT As String
    Dim IsAcct, IsDocs, IsOPs, IsSynd, IsTMT As Long
    Dim RepeatAcct, RepeatDocs, RepeatOps, RepeatSynd, RepeatTMT As Integer
    Dim TimeAcc, TimeDocs, TimeOps, TimeSynd, TimeTMT As Integer
    Dim AcctUser, DocsUser, OpsUser, SyndUser, TMTUser As String

    AcctUser = DLookup("PrimaryContactid", "DepartmentContacts", "Department='Accounting'")
    DocsUser = DLookup("PrimaryContactid", "DepartmentContacts", "Department='Documentation'")
    OpsUser = DLookup("PrimaryContactid", "DepartmentContacts", "Department='Operations'")
    SyndUser = DLookup("PrimaryContactid", "DepartmentContacts", "Department='Syndication'")
    TMTUser = DLookup("PrimaryContactid", "DepartmentContacts", "Department='TMT'")

    'Acctx Rules

    'DocsX Rules
    If Nz(DLookup("SGrpRALB", "schgrp", "sGrpID=" & GroupIDx), 0) = 0 Then DocsX = DocsX & "<li>Group/<b>RAL Returned</b> (date)</li>"
    If Nz(DSum("UnitIFAmt", "Units", "GroupRef=" & GroupIDx), 0) <> 0 And Nz(DMax("UnitPayOff", "Units", "GroupRef=" & GroupIDx), "") = "" _
        Then DocsX = DocsX & "<li>Unit Interim Payoff (date)</li>"
    If Nz(DLookup("SchIn", "Schedule", "SchID=" & DLookup("sGrpSchID", "schgrp", "sgrpID=" & GroupIDx)), "") = "" _
        Then DocsX = DocsX & "<li>Schedule/<b>'Sch-In'</b> (date)</li>"


    'OpsX Rules
    If RST!SGRPNoPD <> True Then
         ' If Nz(DMax("UnitPDStartDate", "PerDiem_Days_Rate_Unit", "GroupRef=" & GroupIDx), "1/1/1900") = "1/1/1900" Then OpsX = OpsX & "<li>Unit PerDiem Start Date</li>"
    End If

    If Nz(DCount("UnitID", "Units", "Nz(UnitSynResid,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx), 0) > 0 Then _
        OpsX = OpsX & "<li>Unit Synd Residual (" & DCount("UnitID", "Units", "Nz(UnitSynResid,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx) & " Units)</li>"

    'SyndX Rules
    If Nz(RST!Assignee_Bank, "") = "" Then SyndX = SyndX & "<li>Group/<b>Assignee Bank</b></li>"
    If Nz(RST!AssigneeBankID, 5469) = 5649 Then SyndX = SyndX & "<li>Group/<b>Assignee Bank</b></li>"
    If Nz(RST!AssignmentDate, "") = "" Then SyndX = SyndX & "<li>Group/<b>Assignment Date</b></li>"
    If Nz(RST!BLCD, "") = "" Then SyndX = SyndX & "<li>Group/<b>BLCD</b></li>"
    If Nz(DLookup("BillToAddressID", "schgrp", "sGrpID=" & GroupIDx), 0) = 0 Then SyndX = SyndX & "<li>Group/Billing Rules/<b>Bill To Address</b></li>"
    If Nz(RST!Fee_Projected, 0) = 0 Then SyndX = SyndX & "<li>Group/<b>Synd Fee</b> (LO)</li>"
    If Nz(RST!SynFee, 0) = 0 Then SyndX = SyndX & "<li>Group/<b>Syndication Fee</b></li>"
    If Nz(RST!SGrpSalesIndex, 9999) = 9999 Then SyndX = SyndX & "<li><b>Group/Index Rate</b> (LO)</li>"
    If Nz(RST!IndexRate, 9999) = 9999 Then SyndX = SyndX & "<li>Group/<b>Index Rate</b></li>"
    If Nz(RST!SgrpSalesYield, 9999) = 9999 Then SyndX = SyndX & "<li>Group/<b>Yield Value</b> (LO)</li>"
    If Nz(RST!SgrpYield, 9999) = 9999 Then SyndX = SyndX & "<li>Group/<b>Yield Value</b></li>"
    If Nz(DLookup("OpsPackage", "SchGrp", "sgrpID=" & GroupIDx), "-") = "-" Then SyndX = SyndX & "<li>Group/<b>OPS Package Date</b></li>"

    If Nz(DCount("UnitID", "Units", "nz(UnitRent,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx), 0) > 0 Then _
        SyndX = SyndX & "<li>Unit <b>Synd</b> Rent (" & DCount("UnitID", "Units", "nz(UnitRent,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx) & " Units)</li>"

    'TMTx Rules
    If Nz(RST!SgrpCommInd, "") = "" Then TMTx = TMTx & "<li>Group/<b>BLCD Basis</b></li>"

    If RST!SGRPNoPD <> True And Nz(DCount("UnitPDRent", "Units", "nz(UnitPDRent,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx), 0) > 0 Then _
        TMTx = TMTx & "<li>Unit PerDiem Rent (" & DCount("UnitPDRent", "Units", "nz(UnitPDRent,0)=0 AND UnitSubType <> 'Auto Hauler' AND GroupRef=" & GroupIDx) & " Units)</li>"

    If Nz(RST!LeaseTermPrimary, 0) = 0 Then TMTx = TMTx & "<li>Group Lease Term Primary</li>"
    If Nz(RST!ExchangeITGroup, 0) <> 0 And Nz(RST!ExchangeITType, "") = "" Then TMTx = TMTx & "<li>Group/<b>ExchangeIT Type</b></li>"
    If Nz(RST!ExchangeITGroup, 0) <> 0 And Nz(RST!ExchangeITMonth, 0) = 0 Then TMTx = TMTx & "<li>Group/<b>ExchangeIT Month</b></li>"
    If Nz(DSum("GroupOEC", "vw_OEC_Rents_LRF_Cost", "GroupID=" & GroupIDx), 0) < 1 Then TMTx = TMTx & "<li>Unit/<b>Cost</b></li>"
    If Nz(RST!LeaseExtOption1Term, 0) <> 0 And RST!LeaseType <> "Full Payout" Then
        If Nz(DCount("FirstExtRent", "Units", "NZ(FirstExtRent,0) > 0 AND GroupRef=" & GroupIDx), 0) < RST!UnitsInGroup Then TMTx = TMTx & "<li>Unit/<b>Extended Rent</b></li>"
    End If
    'Setup bullet List
    If AcctX <> "" Then AcctX = "<b><font size=2>Accounting:</font></b><BR /><ul style='list-style-type:Square'>" & AcctX & "</ul>"
    If DocsX <> "" Then DocsX = "<b><font size=2>Documentation:</font></b><BR /><ul style='list-style-type:Square'>" & DocsX & "</ul>"
    If OpsX <> "" Then OpsX = "<b><font size=2>Operations:</font></b><BR /><ul style='list-style-type:Square'>" & OpsX & "</ul>"
    If SyndX <> "" Then SyndX = "<b><font size=2>Syndication:</font></b><BR /><ul style='list-style-type:Square'>" & SyndX & "</ul>"
    If TMTx <> "" Then TMTx = "<b><font size=2>TMT:</font></b><BR /><ul style='list-style-type:Square'>" & TMTx & "</ul>"

    'Check If there are Existing request, if so update
    IsAcct = Nz(DLookup("INDX", "GroupFinalizedTracker", "NOT isdate(RequestCompleted) AND Department = 'Accounting' AND GroupID=" & GroupIDx), 0)
    IsDocs = Nz(DLookup("INDX", "GroupFinalizedTracker", "NOT isdate(RequestCompleted) AND Department = 'Documentation' AND GroupID=" & GroupIDx), 0)
    IsOPs = Nz(DLookup("INDX", "GroupFinalizedTracker", "NOT isdate(RequestCompleted) AND Department = 'Operations' AND GroupID=" & GroupIDx), 0)
    IsSynd = Nz(DLookup("INDX", "GroupFinalizedTracker", "NOT isdate(RequestCompleted) AND Department = 'Syndication' AND GroupID=" & GroupIDx), 0)
    IsTMT = Nz(DLookup("INDX", "GroupFinalizedTracker", "NOT isdate(RequestCompleted) AND Department = 'TMT' AND GroupID=" & GroupIDx), 0)

    DoCmd.SetWarnings False

    If IsAcct > 0 Then
        Set GRT = CurrentDb.OpenRecordset("Select * From GroupFinalizedTracker Where INDX=" & IsAcct & "", dbOpenSnapshot)
        If IsDate(GRT!Reminderlast) Then TimeAcct = DateDiff("h", GRT!Reminderlast, Now()) Else TimeAcct = DateDiff("h", GRT!OriginalRequest, Now())
        RepeatAcct = GRT!ReminderCount + 1
        If AcctX = "" Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsAcct & ";")
        If AcctX <> "" And TimeAcct >= 11 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set Reminderlast ='" & Now() & "', Reason=""" & AcctX & """, ReminderCount= " & RepeatAcct & "  WHERE INDX=" & IsAcct & ";")
        GRT.close
    End If
    If IsDocs > 0 Then
        Set GRT = CurrentDb.OpenRecordset("Select * From GroupFinalizedTracker Where INDX=" & IsDocs & "", dbOpenSnapshot)
        If IsDate(GRT!Reminderlast) Then TimeDocs = DateDiff("h", GRT!Reminderlast, Now()) Else TimeDocs = DateDiff("h", GRT!OriginalRequest, Now())
        RepeatDocs = GRT!ReminderCount + 1
        If DocsX = "" Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsDocs & ";")
        If DocsX <> "" And TimeDocs >= 11 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set Reminderlast ='" & Now() & "', Reason=""" & DocsX & """, ReminderCount=" & RepeatDocs & " WHERE INDX=" & IsDocs & ";")
        GRT.close
    End If
    If IsOPs > 0 Then
        Set GRT = CurrentDb.OpenRecordset("Select * From GroupFinalizedTracker Where INDX=" & IsOPs & "", dbOpenSnapshot)
        If IsDate(GRT!Reminderlast) Then TimeOps = DateDiff("h", GRT!Reminderlast, Now()) Else TimeOps = DateDiff("h", GRT!OriginalRequest, Now())
        RepeatOps = GRT!ReminderCount + 1
        If OpsX = "" Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsOPs & ";")
        If OpsX <> "" And TimeOps >= 11 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set Reminderlast ='" & Now() & "', Reason=""" & OpsX & """, ReminderCount=" & RepeatOps & " WHERE INDX=" & IsOPs & ";")
        GRT.close
    End If
    If IsSynd > 0 Then
        Set GRT = CurrentDb.OpenRecordset("Select * From GroupFinalizedTracker Where INDX=" & IsSynd & "", dbOpenSnapshot)
        If IsDate(GRT!Reminderlast) Then TimeSynd = DateDiff("h", GRT!Reminderlast, Now()) Else TimeSynd = DateDiff("h", GRT!OriginalRequest, Now())
        RepeatSynd = GRT!ReminderCount + 1
        If SyndX = "" Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsSynd & ";")
        If SyndX <> "" And TimeSynd >= 11 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set Reminderlast ='" & Now() & "', Reason=""" & SyndX & """, ReminderCount=" & RepeatSynd & " WHERE INDX=" & IsSynd & ";")
        GRT.close
    End If
    If IsTMT > 0 Then
        Set GRT = CurrentDb.OpenRecordset("Select * From GroupFinalizedTracker Where INDX=" & IsTMT & "", dbOpenSnapshot)
        If IsDate(GRT!Reminderlast) Then TimeTMT = DateDiff("h", GRT!Reminderlast, Now()) Else TimeTMT = DateDiff("h", GRT!OriginalRequest, Now())
        RepeatTMT = GRT!ReminderCount + 1
        If TMTx = "" Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsTMT & ";")
        If TMTx <> "" And TimeTMT >= 11 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set Reminderlast ='" & Now() & "', Reason=""" & TMTx & """, ReminderCount=" & RepeatTMT & " WHERE INDX=" & IsTMT & ";")
        GRT.close
    End If
    Set GRT = Nothing
    DoCmd.SetWarnings True

    ' Build message for Dialog box
    Msgx = AcctX & DocsX & OpsX & SyndX & TMTx
    If Msgx = "" Then GoTo Cleanup

    Dim ResponseX As String
    Prefix = "<font size=3>Group <b>" & RST!SGrpNO & "</b> Cannot Be Finalized</font><BR />The Fields Below Must Be Corrected:<BR /><br />"
    Suffix = "<BR />Press '<b>Cancel</b>' to cancel Finalization.<br />Press '<b>Send Email</b>' to email the proper teams<br />"
    ' Email without prompt, if selected
    If NoDialog = True Then
        ResponseX = 102
    Else
        If NoEmail = True Then _
            ResponseX = Dialog.RichBox(Prefix & Msgx & Suffix, vbCustom + vbCritical, "Cannot Finalize Group Syndicaton!", , , 0, 1, 1, 0, "Cancel") _
            Else ResponseX = Dialog.RichBox(Prefix & Msgx & Suffix, vbCustom + vbCritical, "Cannot Finalize Group Syndicaton!", , , 0, 1, 1, 0, "Cancel", "Send Email")
    End If
    GroupFinalizedCheck = ResponseX     '101=Cancel
    If ResponseX = 102 Then             '102=Send Email
        Dim Emailx As String
        Dim ContactEmailX As String
        Dim CcEmailX As String
        Dim SubjectX2 As String
        Dim subjectx As String
        Dim CcEmailX2 As String
        Dim Salutation As String
        Dim Sal1 As String
        Dim ii As Integer
        CcEmailX = DLookup("DepartmentEmail", "DepartmentContacts", "INDX=13")
        CcEmailX = Replace(CcEmailX, ";", "@fleetadvantage.com;")
        subjectx = "Cannot Finalize:<b> " & DLookup("ClientShortName", "vw_sixKeys", "SGrpID=" & GroupIDx) & "</b> Sch:<b> " & DLookup("Schedule", "vw_sixKeys", "SgrpId=" & GroupIDx) & "</b> Grp:<b> " & DLookup("UnitGroup", "vw_sixKeys", "SgrpId=" & GroupIDx) & "</b>"
        SubjectX2 = "Cannot Finalize: " & DLookup("ClientShortName", "vw_sixKeys", "SGrpID=" & GroupIDx) & " Sch: " & DLookup("Schedule", "vw_sixKeys", "SgrpId=" & GroupIDx) & " Grp: " & DLookup("UnitGroup", "vw_sixKeys", "SgrpId=" & GroupIDx)
        Prefix = "<p style='color:Navy;font-family:Segoe UI;font-size:14pt;'>" & subjectx & "</p><p style='color:Black;font-family:Segoe UI;font-size:10pt;'>"
        DearAcct = DLookup("[Employee Name]", "TBL_Employees", "ID=" & DLookup("PrimaryContactid", "DepartmentContacts", "Department='Accounting'"))
        DearDocs = DLookup("[Employee Name]", "TBL_Employees", "ID=" & DLookup("PrimaryContactid", "DepartmentContacts", "Department='Documentation'"))
        DearOps = DLookup("[Employee Name]", "TBL_Employees", "ID=" & DLookup("PrimaryContactid", "DepartmentContacts", "Department='Operations'"))
        DearSynd = DLookup("[Employee Name]", "TBL_Employees", "ID=" & DLookup("PrimaryContactid", "DepartmentContacts", "Department='Syndication'"))
        DearTMT = DLookup("[Employee Name]", "TBL_Employees", "ID=" & DLookup("PrimaryContactid", "DepartmentContacts", "Department='TMT'"))
        Sal1 = "Date: " & Date & "<br /><br />"
        Sal1 = Sal1 & "Dear Person,<br /><br />"
        Sal1 = Sal1 & "Please review the list of fields / items <i> shown below</i>, that need <b>your</b> attention and or correction.<br />"
        Sal1 = Sal1 & "<b>Note:</b> A Group Syndication is pending, please attend to this request <b>ASAP</b>"
        Suffix = "</Font><br /> Regards,</p>"

        DoCmd.SetWarnings False

        If AcctX & "test" <> "test" Then
            ContactEmailX = DLookup("[User ID]", "TBL_Employees", "ID=" & AcctUser) & "@fleetadvantage.com"
            CcEmailX2 = Replace(CcEmailX, ContactEmailX & ";", "")
            Salutation = Replace(Sal1, "Person", DearAcct)
            If RepeatAcct > 1 Then Salutation = Salutation & " -This is reminder #<b>" & RepeatAcct & "</b><br /><br />" Else Salutation = Salutation & "<br /><br />"
            If RepeatAcct > 1 Then DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted='" & Now() & "' WHERE INDX=" & IsTMT & ";")
            If IsAcct = 0 Or TimeAcct > 10 Then Call SendEmailMessage(ContactEmailX, SubjectX2, Prefix & Salutation & AcctX & Suffix, True, True, False, CcEmailX2, "", "")
            If IsAcct = 0 Then DoCmd.RunSQL ("INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount) " & _
                "SELECT " & GroupIDx & "," & AcctUser & ",""" & AcctX & """,'Accounting','" & Now & "',1")
        End If
        If DocsX & "test" <> "test" Then
            ContactEmailX = DLookup("[User ID]", "TBL_Employees", "ID=" & DocsUser) & "@fleetadvantage.com"
            CcEmailX2 = Replace(CcEmailX, ContactEmailX & ";", "")
            Salutation = Replace(Sal1, "Person", DearDocs)
            If RepeatDocs > 1 Then Salutation = Salutation & " -This is reminder #<b>" & RepeatDocs & "</b><br /><br />" Else Salutation = Salutation & "<br /><br />"
            If IsDocs = 0 Or TimeDocs > 10 Then Call SendEmailMessage(ContactEmailX, SubjectX2, Prefix & Salutation & DocsX & Suffix, True, True, False, CcEmailX2, "", "")
            If IsDocs = 0 Then DoCmd.RunSQL ("INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount) " & _
                "SELECT " & GroupIDx & "," & DocsUser & ",""" & DocsX & """,'Documentation','" & Now() & "',1")
        End If
        If OpsX & "test" <> "test" Then
            ContactEmailX = DLookup("[User ID]", "TBL_Employees", "ID=" & OpsUser) & "@fleetadvantage.com"
            CcEmailX2 = Replace(CcEmailX, ContactEmailX & ";", "")
            Salutation = Replace(Sal1, "Person", DearOps)
            If RepeatOps > 1 Then Salutation = Salutation & " -This is reminder #<b>" & RepeatOps & "</b><br /><br />" Else Salutation = Salutation & "<br /><br />"
            If IsOPs = 0 Or TimeOps > 10 Then Call SendEmailMessage(ContactEmailX, SubjectX2, Prefix & Salutation & OpsX & Suffix, True, True, False, CcEmailX2, "", "")
            If IsOPs = 0 Then DoCmd.RunSQL ("INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount) " & _
                "SELECT " & GroupIDx & "," & OpsUser & ",""" & OpsX & """,'Operations','" & Now() & "',1")
        End If
        If SyndX & "test" <> "test" Then
            ContactEmailX = DLookup("[User ID]", "TBL_Employees", "ID=" & SyndUser) & "@fleetadvantage.com"
            CcEmailX2 = Replace(CcEmailX, ContactEmailX & ";", "")
            Salutation = Replace(Sal1, "Person", DearSynd)
            If RepeatSynd > 1 Then Salutation = Salutation & " -This is reminder #<b>" & RepeatSynd & "</b><br /><br />" Else Salutation = Salutation & "<br /><br />"
            If IsSynd = 0 Or TimeSynd > 10 Then Call SendEmailMessage(ContactEmailX, SubjectX2, Prefix & Salutation & SyndX & Suffix, True, True, False, CcEmailX2, "", "")
            If IsSynd = 0 Then DoCmd.RunSQL ("INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount) " & _
                "SELECT " & GroupIDx & "," & SyndUser & ",""" & SyndX & """,'Syndication','" & Now() & "',1")
        End If
        If TMTx & "test" <> "test" Then
            ContactEmailX = DLookup("[User ID]", "TBL_Employees", "ID=" & TMTUser) & "@fleetadvantage.com"
            CcEmailX2 = Replace(CcEmailX, ContactEmailX & ";", "")
            Salutation = Replace(Sal1, "Person", DearTMT)
            If RepeatTMT > 1 Then Salutation = Salutation & " -This is reminder #<b>" & RepeatTMT & "</b><br /><br />" Else Salutation = Salutation & "<br /><br />"
            If IsTMT = 0 Or TimeTMT > 10 Then Call SendEmailMessage(ContactEmailX, SubjectX2, Prefix & Salutation & TMTx & Suffix, True, True, False, CcEmailX2, "", "")
            If IsTMT = 0 Then DoCmd.RunSQL ("INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount) " & _
                "SELECT " & GroupIDx & "," & TMTUser & ",""" & TMTx & """,'TMT','" & Now() & "',1")
        End If
        ContactEmailX = Replace(ContactEmailX, "'", "''")
        Emailx = Replace("RE: " & subjectx & "<Br />" & Salutation & Msgx, "'", "''")
        SubjectX2 = "Sign Off: " & Replace(SubjectX2, "'", "''")
        'ContactEmailX = "ljohnson@fleetadvantage.com"
        'CcEmailX2 = "ljohnson@fleetadvantage.com"
        If FaUserName = "Ytj" Then
            Call SendServerEmail(ContactEmailX, CcEmailX2, SubjectX2, Emailx, "yjackson@fleetadvantage.com")
        Else
            Call SendServerEmail(ContactEmailX, CcEmailX2, SubjectX2, Emailx, FaUserName & "@fleetadvantage.com")
        End If
    End If

Cleanup:
    If NoEmail = True Then GroupFinalizedCheck = "check"
    DoCmd.SetWarnings True
    RST.close
    Set RST = Nothing
End Function
Public Function StartSignOff(GroupIDx As Long, SignOffLevel As String)
'Check if the user is allowed to sign-off
    If InStr(1, DLookup("DepartmentEmail", "DepartmentContacts", "SignOffLevel='" & SignOffLevel & "'"), FaUserName) = 0 Then Exit Function

    Dim PrevSigner As Integer
    Dim CurrLevel As Integer
    Dim NextLevel As Integer
    Dim TempStr, YouCanSignOff As String
    Dim ResposeX As VbMsgBoxResultEx

    PrevLevel = CInt(Nz(Right(DMax("Department", "GroupFinalizedTracker", "Department Like 'SignOff*' AND GroupID=" & GroupIDx), 1), "0"))
    CurrLevel = CInt(Right(SignOffLevel, 1))
    NextLevel = NextSignOffLevel(GroupIDx) + 1

    TempStr = CheckIfAlreadySignedOff(GroupIDx, SignOffLevel)
    If TempStr <> "1/1/1900" Then
        ResposeX = Dialog.RichBox("You've Already Signed Off On This Group<br />On: " & TempStr, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel")
        Exit Function
    End If

    YouCanSignOff = "Yes"
    If NextLevel = 0 And CurrLevel = 2 Then YouCanSignOff = "<b>Unit Specs</b> (PO) <b>are confirmed</b>"
    If NextLevel < 3 And CurrLevel = 3 Then YouCanSignOff = "<b>Unit Specs</b> (PO) <b>are confirmed</b><br/><b>Invoices</b> are <b>in and validated</b>"
    If NextLevel < 3 And CurrLevel > 3 Then YouCanSignOff = "<b>Unit Specs</b> (PO) <b>are confirmed</b><br/><b>Invoices</b> are <b>in and validated</b><br/>and the<b> MSO/VINs</b> have been<b> validated</b>"
    If NextLevel <= CurrLevel Then
       TempStr = "Level: " & PrevLevel + 1 & "-'" & DLookup("Department", "DepartmentContacts", "SignOffLevel='SignOff" & PrevLevel + 1 & "'") & "' "
       YouCanSignOff = "The <i><b> Previous Sign-Off</b></i> has been completed<br/><br/> Which is:<b> " & TempStr & "</b>"
    End If
    If YouCanSignOff <> "Yes" Then
        ResposeX = Dialog.RichBox("You cannot sign off unitil:<br/>" & YouCanSignOff, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel")
        Exit Function
    End If

    Dim FromEmail As String
    Dim ToEmail As String
    Dim CcEmail As String
    Dim EmailSuffix As String
    Dim ToPerson As String
    Dim msg1   As String
    Dim ReplyX As String
    Dim StyleX As String
    Dim SQLz As String
    Dim SignOffDesc As String
    Dim SignOffDesc2 As String

    SignOffDesc = "Level: " & CurrLevel & "-'" & DLookup("Department", "DepartmentContacts", "SignOffLevel='" & SignOffLevel & "'") & "' "

    ReplyX = GroupSignOffCheck(GroupIDx, SignOffLevel)      'Check if there are items not completed
    If ReplyX Like "*has detected*" Then StyleX = vbCustom + vbExclamation Else StyleX = vbCustom + vbInformation

    msg1 = msg1 & "<font size=3, Color=Navy>Sign-Off <b>" & SignOffDesc & "</b><br />"
    msg1 = msg1 & "Client:<b> " & DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & Nz(GroupIDx, 0)) & "</b><br/>"
    msg1 = msg1 & "Sch:<b> " & DLookup("Schedule", "vw_SixKeys", "SGrpID=" & Nz(GroupIDx, 0)) & "</b> "
    msg1 = msg1 & "Grp:<b> " & DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & Nz(GroupIDx, 0)) & "</b></font><br/><br/>"
    msg1 = msg1 & ReplyX & ""

    If ReplyX Like "*has detected*" Then
       msg1 = msg1 & "The 'Sign-Off'<b> cannot</b> proceed,<br/>until these items have been corrected<br/>"
       msg1 = msg1 & "<br/><br/>Please press the 'Cancel' button<br/>"
       ResposeX = Dialog.RichBox(msg1, StyleX, AtlaasVersion, , , 0, 1, 0, 0, "Cancel")
    Else
       msg1 = msg1 & "A 'Sign-Off' notification will be emailed<br/>"
       msg1 = msg1 & "<br/><br/>Do you still want to Sign Off?<br/>"
       ResposeX = Dialog.RichBox(msg1, StyleX, AtlaasVersion, , , 0, 1, 0, 0, "Sign-Off", "Cancel")
    End If
    If ResposeX = vbBt2 Or ReplyX Like "*has detected*" Then Exit Function
    If ResposeX = vbBt1 Then

        msg1 = Replace(msg1, "is signing off", "signed off")
        'Post the current sign-off
 '       If NextLevel < 8 Then
            DoCmd.SetWarnings False
            SQLz = "INSERT into GroupFinalizedTracker (GroupID, UserID,Reason,Department,OriginalRequest, ReminderCount,RequestCompleted) " & _
                          "SELECT " & GroupIDx & "," & FAUserIDx & ",""" & msg1 & """,'" & SignOffLevel & "','" & Now & "',1,'" & Now & "'"
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
  '      End If

       ' NextLevel = NextSignOffLevel(GroupIDx)
        EmailSuffix = "@fleetadvantage.com"
        FromEmail = FaUserName & EmailSuffix


        If NextLevel > 7 Then
            SignOffDesc2 = "Sign Off Levels Complete"
            CcEmail = DLookup("DepartmentEmail", "DepartmentContacts", "SignOffLevel='SignOff1'")
            ToPerson = "Laurie Redmond"
            ToEmail = "lredmond;"
        Else
            If NextLevel <= CurrLevel Then NextLevel = PrevLevel + 1 Else NextLevel = CurrLevel + 1

            SignOffDesc2 = "Level " & NextLevel & "-'" & DLookup("Department", "DepartmentContacts", "SignOffLevel='SignOff" & NextLevel & "'") & "' "
            ToPerson = DLookup("[Employee Name]", "TBL_Employees", "ID =" & DLookup("PrimaryContactID", "DepartmentContacts", "SignOffLevel='SignOff" & NextLevel & "'"))
            ToEmail = DLookup("[User ID]", "TBL_Employees", "ID =" & DLookup("PrimaryContactID", "DepartmentContacts", "SignOffLevel='SignOff" & NextLevel & "'")) & ";"
            CcEmail = DLookup("DepartmentEmail", "DepartmentContacts", "SignOffLevel='SignOff" & CurrLevel & "'")
        End If

        'CcEmail = Replace(CcEmail, FromEmail, "")               'Remove the From Email
        'CcEmail = Replace(CcEmail, ToEmail, "")                 'Remove the To Email
        CcEmail = Replace(CcEmail, ";", (EmailSuffix & ";"))    'Add in the @address.com;
        CcEmail = Left(CcEmail, Len(Trim(CcEmail)) - 1)         'Take off the trailing ';'
        ToEmail = Replace(ToEmail, ";", EmailSuffix)
        Call SubmitSignOff(GroupIDx, SignOffLevel, SignOffDesc, SignOffDesc2, SplitString(msg1, "A 'Sign-Off'", 0), ToEmail, ToPerson, CcEmail, CurrLevel, NextLevel)
    End If
End Function
Public Function CheckIfAlreadySignedOff(GroupIDx As Long, SignOffLevel As String) As Date
    Dim LevelSignedOff As Date
    CheckIfAlreadySignedOff = Nz(DLookup("OriginalRequest", "GroupFinalizedTracker", "Department = '" & SignOffLevel & "' AND GroupID=" & GroupIDx), "1/1/1900")
End Function
Public Function NextSignOffLevel(GroupIDx As Long) As Integer
    Dim NextLevel, Lx, Nx As Integer
    On Error Resume Next
    NextLevel = CInt(Nz(Right(DMax("Department", "GroupFinalizedTracker", "GroupID=" & GroupIDx & " AND Department Like 'SignOff*'"), 1), 0))
FixMe:
    On Error GoTo 0
    NextSignOffLevel = NextLevel + 1
End Function

Public Function SubmitSignOff(GroupIDx As Long, SignOffLevel As String, SignOffDesc As String, SignOffDesc2 As String _
    , ByRef msg1 As String _
    , ByRef ToEmail As String _
    , ByRef ToPerson As String, CcEmailX As String, CurrLevel, NextLevel As Integer)
    'Update Previous sign off
    If CurrLevel > 1 Then
        DoCmd.SetWarnings False
        '      DoCmd.RunSQL ("Update GroupFinalizedTracker Set RequestCompleted ='" & Date & "' " & _
               "WHERE GroupId=" & GroupIDx & " AND Department='" & "SignOff" & CInt(Right(SignOffLevel, 1)) - 1 & "'")
        DoCmd.SetWarnings True
    End If

    'Send Email
    Dim Emailx As String
    Dim Subject2X, body, TargetGroup, TargetGroup2 As String
    Dim subjectx As String
    msg1 = Replace(msg1, "<font size=3,", "<font size=2,")
    '' Dim ii As Integer
    TargetGroup = DLookup("ClientShortName", "vw_sixKeys", "SGrpID=" & GroupIDx) & ", Sch: " & DLookup("Schedule", "vw_sixKeys", "SgrpId=" & GroupIDx) & " Grp: " & DLookup("UnitGroup", "vw_sixKeys", "SgrpId=" & GroupIDx) & ""
    TargetGroup2 = "<b> " & DLookup("ClientShortName", "vw_sixKeys", "SGrpID=" & GroupIDx) & "</b> Sch:<b> " & DLookup("Schedule", "vw_sixKeys", "SgrpId=" & GroupIDx) & "</b> Grp:<b> " & DLookup("UnitGroup", "vw_sixKeys", "SgrpId=" & GroupIDx) & "</b>"

    subjectx = "Sign Off: " & TargetGroup & " - " & SignOffDesc & " was submitted"

    If Right(SignOffLevel, 1) = "1" Then
        Subject2X = "Sign Off: " & TargetGroup2 & " process has started"
    Else
        Subject2X = "Sign Off: " & TargetGroup2 & " Notice"
    End If

    Prefix = "<p style='color:Navy;font-family:Segoe UI;font-size:14pt;'>" & Subject2X & "</p><p style='color:Black;font-family:Segoe UI;font-size:10pt;'>"

    body = "Date: " & Date & " From: " & FAUserFullName & "<br /><br />"
    body = body & "Dear " & ToPerson & ",<br /><br />"
    If NextLevel < 8 Then
        body = body & "You've been assigned this Group for Sign-Off <b>" & SignOffDesc2 & "</b><br/>"
        body = body & "Please attend to your section of the sign-off process, You will be confirming the following:"
        body = body & DLookup("SignOffTaskList", "DepartmentContacts", "SignOffLevel='" & "SignOff" & NextLevel & "'") & " <br/>"
        body = body & "Regards,<br/><br/>"
        body = body & "<hr>For reference, the <i><b> previous</b></i> 'Group Sign Offs', shown below:<hr><br/>"
    End If
    If NextLevel > 1 Then
        Dim PrevSignOffs
        Dim RST As DAO.Recordset
        Set RST = CurrentDb.OpenRecordset("SELECT E.[Employee Name], DC.Department, DC.SignOffTaskList, T.GroupID, T.OriginalRequest " & _
            "FROM (DepartmentContacts As DC INNER JOIN GroupFinalizedTracker as T ON DC.SignOffLevel = T.Department) " & _
            "INNER JOIN TBL_Employees As E ON T.UserID = E.ID " & _
            "WHERE T.groupID = " & GroupIDx & " AND DC.SignOffLevel Like 'SignOff*'ORDER BY DC.INDX;", dbOpenSnapshot)
        RST.MoveLast
        RST.MoveFirst
        Do While Not RST.EOF
            PrevSignOffs = PrevSignOffs & "<p><b>" & RST!Department & "</b><br/>"
            PrevSignOffs = PrevSignOffs & "by: " & RST![Employee Name] & " On: " & RST!OriginalRequest & "</p>"
            PrevSignOffs = PrevSignOffs & RST!signofftaskList
            PrevSignOffs = PrevSignOffs & ""
            RST.MoveNext
        Loop
        RST.close
        Set RST = Nothing
    End If
    If NextLevel = 1 Then Emailx = Prefix + body + msg1 & "</p>" Else Emailx = Prefix + body + PrevSignOffs & "</p>"
    If NextLevel > 7 Then
        body = body & "The Group<i><b> 'Sign-Off'</b></i> Process for the above mentioned Group, is now <b> complete</b>.<br/>"
        body = body & "If there are any '<u>special</u>' or '<u>outstanding</u>' requirements needed for this Group,<br/>please contact the proper Team/Department as needed:<br/><br/>"
        body = body & "Regards,<br/><br/>"
        body = body & "<hr>For reference, the <i><b> previous</b></i> 'Group Sign Offs', are shown below:<hr><br/>"
        Emailx = Prefix + body + PrevSignOffs + ReplyX & "</p>"
    End If
    ReplyX = GroupSignOffCheck(GroupIDx, "SignOff8")
    'ToEmail = "ljohnson@fleetadvantage.com"
    'CcEmailX = "ljohnson@fleetadvantage.com"
    '********************************************************
    '*  Replaced Outlook email to SQL email 10/14/2022 Ljohnson
    '*  Call SendEmailMessage(ToEmail, subjectx, EmailX, False, True, False, CcEmailX, "", "")
    '********************************************************
    Emailx = Replace(Emailx, "'", "''")
    subjectx = Replace(subjectx, "'", "''")
    If FaUserName = "ytj" Then
        Call SendServerEmail(ToEmail, CcEmailX, subjectx, Emailx, "yjackson@fleetadvantage.com")
    Else
        Call SendServerEmail(ToEmail, CcEmailX, subjectx, Emailx, FaUserName & "@fleetadvantage.com")
    End If
End Function

Public Function GroupSignOffCheck(GroupIDx As Long, SignOffLevel As String, Optional NoDialog As Boolean, Optional NoEmail As Boolean) As String
    Dim RST    As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset("Select * From vw_UnitGroupHealth Where GroupID=" & GroupIDx, dbOpenSnapshot)
    Dim SQLz   As String
    Dim POSet, SpecSet As Integer
    Dim SignOffTask, TaskMsg, Testmsg As String


    GroupSignOffCheck = "Good"
    TaskMsg = ""
    If SignOffLevel <> "SignOff8" Then
        SignOffTask = FAUserFullName & " is signing off, indicating that:<br/><ul>" & Nz(DLookup("SignoffTaskList", "DepartmentContacts", "SignOffLevel='" & SignOffLevel & "'"), "0") & "</ul><br/>"
    End If
    Select Case SignOffLevel
    Case "SignOff1"
        If RST!Invoiced <> "Y" Then TaskMsg = TaskMsg & "<li>Invoices Missing: " & RST!Invoiced & " entered </li>"

    Case "SignOff2" 'KB or Nathan Rhoads
        SpecSet = DCount("unitId", "vw_Sixkeys", "UnitDescDocVer like '*tbd*' AND sGrpID=" & GroupIDx & "")
        POSet = Nz(DLookup("POSentDt", "schgrp", "sGrpID=" & GroupIDx), 0)
        If POSet = 0 Then TaskMsg = TaskMsg & "<li>P.O. date not entered</li>"
        If SpecSet > 0 Then TaskMsg = TaskMsg & "<li>Spec Incomplete: " & SpecSet & "</li>"

    Case "SignOff3"
        If RST!MSOSet <> "Y" Then TaskMsg = TaskMsg & "<li>MSOs Missing: " & RST!MSOSet & " entered </li>"

    Case "SignOff3"
        If RST!VINSet <> "Y" Then TaskMsg = TaskMsg & "<li>VINs Missing: " & VINSet & " entered </li>"
        If RST!LDelivery <> "Y" Then TaskMsg = TaskMsg & "<li>Delv Loc Missing: " & RST!LDelivery & " entered </li>"
        If RST!LGarage <> "Y" Then TaskMsg = TaskMsg & "<li>Garage Missing: " & RST!LGarage & " entered </li>"
        If RST!LReg <> "Y" Then TaskMsg = TaskMsg & "<li>Reg Loc Missing: " & RST!LReg & " entered </li>"
        If RST!LTitle <> "Y" Then TaskMsg = TaskMsg & "<li>Reg Loc Missing: " & RST!LTitle & " entered </li>"
        ' If InvoicePaid <> "Y" Then  TaskMsg = TaskMsg & "<li>Open Invoices: " & InvoicePaid & "</li>"

    Case "SignOff5"
        If RST!PDRent <> "Y" Then TaskMsg = TaskMsg & "<li>PD Rent Missing: " & RST!PDRent & " entered </li>"
        If RST!RentSCH <> "Y" Then TaskMsg = TaskMsg & "<li>Sch Rent Missing: " & RST!RentSCH & " entered </li>"
        If RST!RentExtd <> "Y" Then TaskMsg = TaskMsg & "<li>Extended Rent Missing: " & RST!RentExtd & " entered </li>"
        If RST!RVSales <> "Y" Then TaskMsg = TaskMsg & "<li>Sch Residual Missing: " & RST!RVSales & " entered </li>"

    Case "SignOff6"
        If RST![Sch-In] <> "Y" Then TaskMsg = TaskMsg & "<li>Sch-In date Not set</li>"
        If RST!RentSYND <> "Y" Then TaskMsg = TaskMsg & "<li>Synd Rent Missing on: " & RST!RentSYND & " units</li>"
        If RST![RAL-In] <> "Y" Then TaskMsg = TaskMsg & "<li>RAL date Not set</li>"
        If RST!Billto <> "Y" Then TaskMsg = TaskMsg & "<li>Bill To Not set</li>"
        If RST!OPSPkg <> "Y" Then TaskMsg = TaskMsg & "<li>OPS Package Received Date Not set</li>"

    Case "SignOff7"
        If RST!PDStart <> "Y" Then TaskMsg = TaskMsg & "<li>Per-Diem Start Date Missing: " & RST!PDStart & " entered</li>"
        If RST!RentSYND <> "Y" Then TaskMsg = TaskMsg & "<li>Synd Rent Missing on: " & RST!RentSYND & " units</li>"
        If RST!RVSynd <> "Y" Then TaskMsg = TaskMsg & "<li>Synd Resiudal Missing: " & RST!RVSynd & " entered </li>"

    Case "SignOff8"
        If RST!MSOSet <> "Y" Then TaskMsg = TaskMsg & "<li>MSOs Missing: " & RST!MSOSet & " entered </li>"
        If RST!Invoiced <> "Y" Then TaskMsg = TaskMsg & "<li>Invoices Missing: " & RST!Invoiced & " entered </li>"
        If RST!VINSet <> "Y" Then TaskMsg = TaskMsg & "<li>VINs Missing: " & VINSet & " entered </li>"
        If RST!LDelivery <> "Y" Then TaskMsg = TaskMsg & "<li>Delv Loc Missing: " & RST!LDelivery & " entered </li>"
        If RST!LGarage <> "Y" Then TaskMsg = TaskMsg & "<li>Garage Missing: " & RST!LGarage & " entered </li>"
        If RST!LReg <> "Y" Then TaskMsg = TaskMsg & "<li>Reg Loc Missing: " & RST!LReg & " entered </li>"
        If RST!LTitle <> "Y" Then TaskMsg = TaskMsg & "<li>Reg Loc Missing: " & RST!LTitle & " entered </li>"
        SpecSet = DCount("unitId", "vw_Sixkeys", "UnitDescDocVer like '*tbd*' AND sGrpID=" & GroupIDx & "")
        POSet = Nz(DLookup("POSentDt", "schgrp", "sGrpID=" & GroupIDx), 0)
        If POSet = 0 Then TaskMsg = TaskMsg & "<li>P.O. date not entered</li>"
        If SpecSet > 0 Then TaskMsg = TaskMsg & "<li>Spec Incomplete: " & SpecSet & "</li>"
        If RST!PDRent <> "Y" Then TaskMsg = TaskMsg & "<li>PD Rent Missing: " & RST!PDRent & " entered </li>"
        If RST!RentSCH <> "Y" Then TaskMsg = TaskMsg & "<li>Sch Rent Missing: " & RST!RentSCH & " entered </li>"
        If RST!RentExtd <> "Y" Then TaskMsg = TaskMsg & "<li>Extended Rent Missing: " & RST!RentExtd & " entered </li>"
        If RST!RVSales <> "Y" Then TaskMsg = TaskMsg & "<li>Sch Residual Missing: " & RST!RVSales & " entered </li>"
        If RST![Sch-In] <> "Y" Then TaskMsg = TaskMsg & "<li>Sch-In date Not set</li>"
        If RST![RAL-In] <> "Y" Then TaskMsg = TaskMsg & "<li>RAL date Not set</li>"
        If RST!Billto <> "Y" Then TaskMsg = TaskMsg & "<li>Bill To Not set</li>"
        If RST!OPSPkg <> "Y" Then TaskMsg = TaskMsg & "<li>OPS Package Received Date Not set</li>"
        If RST!PDStart <> "Y" Then TaskMsg = TaskMsg & "<li>Per-Diem Start Date Missing on : " & RST!PDStart & " units</li>"
        If RST!RentSYND <> "Y" Then TaskMsg = TaskMsg & "<li>Synd Rent Missing on: " & RST!RentSYND & " units</li>"
        If RST!RVSynd <> "Y" Then TaskMsg = TaskMsg & "<li>Synd Resiudal Missing on: " & RST!RVSynd & " units </li>"
    End Select
    RST.close
    Set RST = Nothing
Cleanup:
    DoCmd.SetWarnings True
    If TaskMsg = "" Then
        TaskMsg = "<b>ATLAAS</b> has NOT detected any issues.<br/><br/>"
    Else
        TaskMsg = "<b>ATLAAS has detected the following:</b><br/><ul>" & TaskMsg & "</ul><br/>"
    End If
    GroupSignOffCheck = SignOffTask & TaskMsg
End Function
