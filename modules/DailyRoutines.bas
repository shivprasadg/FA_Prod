Attribute VB_Name = "DailyRoutines"
Option Compare Database

Public SQLz As String
Public SQLWhere As String
Public subjectx  As String
Public ContactEmailX As String
Public CcEmailX As String
Public msg1 As String
Public Linkx As String
Public XLS As String
Public XLSpre As String



Public Function SetReportStartEnd()
DoCmd.Requery "ReportStartEnd_MakeTemp"
DoCmd.Requery "ReportStartEnd_ABUnits_Update"
DoCmd.Requery "ReportStartEnd_TRUnits_Update"
DoCmd.Requery "ReportStartEnd_SKUnits_Update"
End Function

Public Function SendOverMilesEmail()
    'Source SQL for Both Tabs
    SQLz = "SELECT "
    SQLz = SQLz & " Client"                     'col A
    SQLz = SQLz & ",Sch"                        'col B
    SQLz = SQLz & ",Grp"                        'col C
    SQLz = SQLz & ",AssetID"                    'col D
    SQLz = SQLz & ",Unitnum"                    'col E
    SQLz = SQLz & ",VIN"                        'col F
    SQLz = SQLz & ",MYear"                      'col G
    SQLz = SQLz & ",Make"                       'col H
    SQLz = SQLz & ",Model"                      'col I
    SQLz = SQLz & ",[Type]"                     'col J
    SQLz = SQLz & ",SubType"                    'col K
    SQLz = SQLz & ",Domicile"                   'col L
    SQLz = SQLz & ",BLCD"                       'col M
    SQLz = SQLz & ",[Lease Term]"               'col N
    SQLz = SQLz & ",[Extd Mths]"                'col O
    SQLz = SQLz & ",[Lease Age]"                'col P
    SQLz = SQLz & ",[Mths Left]"                'col Q
    SQLz = SQLz & ",[Lease Expire]"             'col R
    SQLz = SQLz & ",[Off Lease Projected]"      'col S
    SQLz = SQLz & ",[Odo Date]"                 'col T
    SQLz = SQLz & ",[ODO Read]"                 'col U
    SQLz = SQLz & ",[Odo Source]"               'col V
    SQLz = SQLz & ",[MPM Allowed]"              'col W
    SQLz = SQLz & ",[MPM Actual]"               'col X
    SQLz = SQLz & ",[MPY Allowed]"              'col Y
    SQLz = SQLz & ",[MPY  Actual]"              'col Z
    SQLz = SQLz & ",0 as [Miles Allowed (today)]"                       'col AA

    SQLz = SQLz & ",[Odo Read] as [Miles Actual (today)]"           'col New AB

    SQLz = SQLz & ",[Pct Used] AS [Pct Actual]"                         'col AC

    SQLz = SQLz & ",[Miles OvrUndr] AS [Miles OvrUndr (today)]"         'col AD

    SQLz = SQLz & ",[Over Miles Charge] as [Over Miles Charge (today)]" 'col AE
    SQLz = SQLz & ",[Total Miles Allowed] AS [Miles Allowed proj'd]"    'col AF
    SQLz = SQLz & ",ODOProj as [Miles Used proj'd]"                     'col AG
    SQLz = SQLz & ",OverUndrProj  AS [Miles OvrUndr proj'd]"            'col AH
    SQLz = SQLz & ",OverMilesProj AS [Over Miles Charge proj'd]"        'col AI
    SQLz = SQLz & " FROM vw_OverMilesEmail"

    Linkx = "\\fa-poseidon\Atlaas\Temp\Mileage_c__" & DatePart("M", Date) & "_" & DatePart("D", Date) & "_" & DatePart("YYYY", Date) & ".xlsx"
    Call KillFiles("\\fa-poseidon\Atlaas\Temp", "ExpiringUnits*.xlsx")
   'XLS Script for BOTH Tabs
    XLS = ""
    XLS = XLS & "|A,DD3|A,TL"
    XLS = XLS & "|B,TC"
    XLS = XLS & "|D,BX|D,TC|D,cColor:3:black|D,cBorder:Med"
    XLS = XLS & "|E,TC"
    XLS = XLS & "|G,TC|G,TC"
    XLS = XLS & "|M,TC|M,W10|M,cColor:3:black|M,cBorder:Med"
    XLS = XLS & "|N,TC|N,W6"
    XLS = XLS & "|O,TC|O,W6"
    XLS = XLS & "|P,TC|P,W6"
    XLS = XLS & "|Q,TC|Q,W6"
    XLS = XLS & "|R,TC|R,W10"
    XLS = XLS & "|S,TC|S,W10"
    XLS = XLS & "|T,TC|T,W10"
    XLS = XLS & "|U,Sint|U,W8"
    XLS = XLS & "|V,BC|V,W7|V,cBorder:Med"
    XLS = XLS & "|W,cLabel:End=AC:2:" & TLeftLine & LineSpace(20) & " Mileage Allowance vs Usage " & LineSpace(20) & TRightLine
    XLS = XLS & "|W,cFontColor:2:white|W,Sint|W,W8|W,cColor:3:teal:End=AC|W,cColor:2:teal|W,TC"
    XLS = XLS & "|X,Sint|X,W8"
    XLS = XLS & "|Y,Sint|Y,W8"
    XLS = XLS & "|Z,Sint|Z,W9"
    XLS = XLS & "|AA,FxAll=[@[MPM Allowed]]*[@[Lease Age]]|AA,Sint|AA,W10"
    XLS = XLS & "|AB,Sint|AB,W9"
    XLS = XLS & "|AC,TC|AC,%|AC,W6|AC,cColor:3:black|AC,cBorder:Med"
    XLS = XLS & "|AD,Sint|AD,cColor:3:orange:End=AE|AD,W9"
    XLS = XLS & "|AE,$E|AE,W11|AE,cBorder:Med"
    XLS = XLS & "|AF,Sint|AF,W10"
    XLS = XLS & "|AF,cLabel:End=AH:2:" & TLeftLine & LineSpace(5) & " Projected Miles At Term " & LineSpace(5) & TRightLine
    XLS = XLS & "|AF,cColor:3:dark green:End=AH|AF,cColor:2:dark green|AF,cFontColor:2:white"
    XLS = XLS & "|AG,Sint|AG,W10"
    XLS = XLS & "|AH,Sint|AH,W10"
    XLS = XLS & "|AI,$E|AI,W13|AI,cColor:3:brown|AI,cBorder:Med"
    XLS = XLS & "|"

    'Over Miles tab 1
    Call MessageUser("OPEN", "Making the OVER Miles Tab", "Excel will now open...")

    SQLWhere = " WHERE [Pct Used] > 1.24  ORDER BY Client, [Pct Used] DESC"
    XLSpre = "|A,mTable::Name=Over"
    SendReport2Excel SQLz & SQLWhere, "OVER Miles", "Units 125% and Over Utilization: " & Date, True, XLSpre & XLS, 0, False, 1, False, , Linkx

    'Under Miles tab 2
    Call MessageUser("Update", "Making the UNDER Miles Tab", "Four more tabs...")

    SQLWhere = " WHERE [Pct Used] < 0.61 ORDER BY Client, [Pct Used] "
    XLSpre = "|A,mTable:Style=Medium3:Name=Under"
    SendReport2Excel SQLz & SQLWhere, "UNDER Miles", "Units 60% and Under Utilization: " & Date, True, XLSpre & XLS, 0, False, 2, False, , Linkx

    Call MessageUser("OPEN", "Sending Email...", "with an attachment, just a bit more")
    subjectx = "Notice: FA Units with Over/Under Utilization"
    'make email message
    ContactEmailX = "jdusseau@FleetAdvantage.com"
    CcEmailX = "TMT@FleetAdvantage.com;ljohnson@FleetAdvantage.com"

    'ContactEmailX = "ljohnson@FleetAdvantage.com" ' for testing
    'CcEmailX = "it@FleetAdvantage.com" ' for testing
    msg1 = ""
    msg1 = msg1 & "<html><body><H2 style=""color:Navy;font-family:Segoe UI;""><b>Notice: FA Units with Over/Under Utilization " & Date & "</H2></b></p>"
    msg1 = msg1 & "<p style=""font-family:Calibri;font-size:11pt"">"
    msg1 = msg1 & "Re: " & subjectx & "<br/><br/>"
    msg1 = msg1 & "Please review the attached spreadsheet: " & SplitString(Linkx, "Temp\", 1) & "<br />"
    msg1 = msg1 & "<b>Note:</b> There are two tabs 1: OVER Miles 2: UNDER Miles<br /><br />"
    msg1 = msg1 & "Thank You For Your Attention In This Matter<br /><br />"
    msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:8px""><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:12px"">"
    msg1 = msg1 & "</body></html>"
    Call SendServerEmail(ContactEmailX, CcEmailX, subjectx, msg1, "IT@fleetadvantage.com", Linkx)
    Call MessageUser("Close")
End Function
