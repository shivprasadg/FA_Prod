Attribute VB_Name = "Paymetadjustments"
Option Compare Database


Public Function MakePaymentAdjustment(lGroupID As Long)
    'Exit Function
    If FAUserAccessGroup Like "*TMT*" Then
        Dim SQLz As String
        Dim Headerx As String
        Dim ReasonX As String
        Dim subjectx As String
        Dim msg1 As String
        Dim i As Long
        Dim RST As DAO.Recordset
        SQLz = " Select Client"                             ' Col A
        SQLz = SQLz & ",Sch"                                ' Col B
        SQLz = SQLz & ",Grp"                                ' Col C
        SQLz = SQLz & ",[Fund Status]"                      ' Col D
        SQLz = SQLz & ",AssetID"                            ' Col E
        SQLz = SQLz & ",[Unit #]"                           ' Col F
        SQLz = SQLz & ",VIN"                                ' Col G
        SQLz = SQLz & ",[LPAF Cost]"                        ' Col H
        SQLz = SQLz & ",[Unit OEC]"                         ' Col I
        SQLz = SQLz & ",UnitCost"                           ' Col J
        SQLz = SQLz & ",PartCost"                           ' Col K
        SQLz = SQLz & ",UpFrontTax"                         ' Col L
        SQLz = SQLz & ",[LPAF Rent] "                       ' Col M
        SQLz = SQLz & ",UNITSRENT"                          ' Col N
        SQLz = SQLz & ",NewSchRent"                         ' Col O
        SQLz = SQLz & ",NewPerDiem"                         ' Col P
        SQLz = SQLz & ",NewExtRent"                         ' Col Q
        SQLz = SQLz & ",[Error] "                           ' Col R
        SQLz = SQLz & ",ATFGroupSpecID AS [Group SpecID]"   ' Col S
        SQLz = SQLz & ",ATFUnitSpecID AS [Unit SpecID]"     ' Col T
        SQLz = SQLz & " FROM vw_PaymentAdj "
        SQLz = SQLz & " WHERE (GroupID=" & lGroupID & " AND SignOfflevel < 5 AND NewSchRent <> UNITSRENT)"

        Set RST = CurrentDb.OpenRecordset(SQLz)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        If RST.RecordCount = 0 Then Exit Function
        Headerx = "Sch Rent was updated for "
        If i > 1 Then Headerx = Headerx & i & " units " Else Headerx = Headerx & i & " unit "
        Headerx = Headerx & "Sch: " & RST!Sch & " Grp: " & RST!Grp & ", as follows (" & Date & "):"
        subjectx = "Cost / Sch Rent / Per Diem Change: " & RST!Client & " | Sch " & RST!Sch & " | Grp " & RST!Grp
        'make email message
        msg1 = msg1 & "<html><body><H3 style=""color:Navy;font-family:Segoe UI;""><b>Cost / Sch Rent / Per Diem Change - " & Date & "</H3></b></p>"
        msg1 = msg1 & "<p style=""font-family:Calibri;font-size:11pt"">"
        msg1 = msg1 & "Re: " & subjectx & "<br/><br/>"
        msg1 = msg1 & "Client:  <B>" & RST!Client & "</B><br />"
        msg1 = msg1 & HSpaces(5) & "Schedule: <B>" & RST!Sch & "</B>"
        msg1 = msg1 & HSpaces(5) & "Group: <B>" & RST!Grp & "</B> "
        msg1 = msg1 & HSpaces(5) & "Total Units changed: <B>" & RST.RecordCount & "</B><Br />"
        msg1 = msg1 & HSpaces(5) & "Date changes made: <B> " & Date & "</B> "
        msg1 = msg1 & HSpaces(5) & "Changes made by: <B>By Atlaas Core</B> (<i>not " & FaUserName & "</i>) :)<Br />"
        msg1 = msg1 & "<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=2>"
        msg1 = msg1 & "<TR ALIGN=""Center"", bgcolor=""Navy"">"
        msg1 = msg1 & "<TH>LPAF Cost</TH>"
        msg1 = msg1 & "<TH>Curr OEC</TH>"
        msg1 = msg1 & "<TH>Rent From</TH>"
        msg1 = msg1 & "<TH>Rent To</TH>"
        msg1 = msg1 & "<TH>*Per Diem</TH>"
        msg1 = msg1 & "<TH>Unit Num</TH>"
        msg1 = msg1 & "<TH>VIN</TH>"
        msg1 = msg1 & "<TH>Reason</TH>"
        msg1 = msg1 & "<TH>*UpFront Tax</TH>"
        msg1 = msg1 & "</TR>"
        RST.MoveFirst
        For i = 1 To RST.RecordCount
            If RST![LPAF Cost] > RST![Unit OEC] Then ReasonX = "Cost Down" Else ReasonX = "Cost Up"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST![LPAF Cost], "$#,#.#0") & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST![Unit OEC], "$#,#.#0") & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST!UNITSRENT, "$#,#.#0") & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST!NewSchRent, "$#,#.#0") & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST!NewPerDiem, "$#,#.#0") & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & RST![Unit #] & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & RST!VIN & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & ReasonX & " </TD>"
            msg1 = msg1 & "<TD ALIGN=""Center"">" & Format(RST!UpFrontTax, "$#,#.#0") & " </TD>"
            msg1 = msg1 & "</span></TR>"
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Update Units Set UnitSRent =" & RST!NewSchRent & ", UNITPDRENT= " & RST!NewPerDiem & ", FirstExtRent= " & RST!NewExtRent & "  WHERE UnitID = " & RST!AssetID & ";")
            DoCmd.SetWarnings True
            RST.MoveNext
        Next i

        msg1 = msg1 & "</table>"
        msg1 = msg1 & "<br/><b>*Per Diem Rent EXCLUDES Up Front Tax</b>"
        msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:8px""><br/>[End Msg] "
        msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
        msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:12px"">"
        msg1 = msg1 & "</body></html>"
        Call SendServerEmail("TMT@fleetadvantage.com;Operations@fleetadvantage.com", "ljohnson@fleetadvantage.com", subjectx, msg1, "TMT@fleetadvantage.com")
        '  Call SendEmailMessage("TMT@fleetadvantage.com", subjectx, msg1, True, True, False, "ljohnson@fleetadvantage.com")
Makespreadsheet:
      '  Call MessageUser("Open", "Getting Payment Adjuments", "...Excel will open soon...")
        XLS = ""
        XLS = "|A,DD3|A,Wrap|B,TC|C,TC|D,TC|E,BC|H,$R|I,$R|J,$R|K,$R|L,$R|M,$R|N,$R|O,$R|P,$R|Q,$R|S,TC|T,TC"
        XLS = XLS & "|A,W15|B,W7|D,W06|N,W10|S,W08|T,W08"
        XLS = XLS & "|H,cColor:3:green|I,cColor:3:purple|L,cColor:3:green|M,cColor:3:purple|N,cColor:3:crimson|O,cColor:3:black"
        XLS = XLS & "|H,cBorder:Med|I,cBorder:Med|L,cBorder:Med|M,cBorder:Med|N,cBorder:Med|"
   '     SendReport2Excel SQLz, "adj", Headerx, True, XLS, 0, False
        RST.close
        Set RST = Nothing

    End If
    Call MessageUser("Close")
End Function
