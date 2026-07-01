Attribute VB_Name = "Excelexporter"
Option Compare Database
Option Explicit
    Dim SQLz As String
    Dim XLS As String
    Dim Headerx As String
Public Function PriceReconReport(ClientIDi As Long, GroupIDi As Long, FAIDx As String)
    Dim SQLx, SchNoX, Grpx As String
    Dim rCntLPAF, rCntCURR As Integer

    XLS = ""
    SchNoX = DLookup("Schedule", "vw_SixKeys", "sGrpid=" & GroupIDi)
    Grpx = DLookup("UnitGroup", "vw_SixKeys", "sGrpid=" & GroupIDi)
    Headerx = "Pricing Reconn for Sch: " & SchNoX & " Grp: " & Grpx & " as of: " & Format(Date, "mm-dd-yy")
    SQLz = "Select "
    SQLz = SQLz & " Grp"                     '-Col A
    SQLz = SQLz & ",FAID"                    '-Col B
    SQLz = SQLz & ",Units"                   '-Col C
    SQLz = SQLz & ",UnitYr"                  '-Col D
    SQLz = SQLz & ",UnitMake"                '-Col E
    SQLz = SQLz & ",UnitModel"               '-Col F
    SQLz = SQLz & ",UnitType"                '-Col G
    SQLz = SQLz & ",UnitSubType"             '-Col H
    SQLz = SQLz & ",Trans"                   '-Col I
    SQLz = SQLz & ",UnitCost"                '-Col J
    SQLz = SQLz & ",PartsCost"               '-Col K
    SQLz = SQLz & ",OEC"                     '-Col L
    SQLz = SQLz & ",RentLO"                  '-Col M
    SQLz = SQLz & ",LeaseType"               '-Col N
    SQLz = SQLz & ",RateTerm"                '-Col O
    SQLz = SQLz & ",RateLO"                  '-Col P
    SQLz = SQLz & ",SwapDate"                '-Col Q
    SQLz = SQLz & ",RVLPAF"                  '-Col R
    SQLz = SQLz & ",LRF"                     '-Col S
    SQLz = SQLz & ",LTerm"                   '-Col T
    SQLz = SQLz & " FROM vw_PricingRecon_LPAF"
    SQLz = SQLz & " WHERE ClientID =" & ClientIDi & " "
    SQLz = SQLz & " AND FAID = '" & FAIDx & "' "
    SQLz = SQLz & " ORDER BY Grp"
    rCntLPAF = DCount("FAID", "vw_PricingRecon_LPAF", "ClientID = " & ClientIDi & " AND FAID = '" & FAIDx & "'")
    XLS = "|A,W15|A,BC|C,TE|J,$E|K,$E|L,$E|M,$E|N,TC|O,TC|R,$R|"

    SQLx = "Select "
    SQLx = SQLx & " Grp"                     '-Col A
    SQLx = SQLx & ",FAID"                    '-Col b
    SQLx = SQLx & ",Units"                   '-Col C
    SQLx = SQLx & ",UnitYr"                  '-Col D
    SQLx = SQLx & ",UnitMake"                '-Col E
    SQLx = SQLx & ",UnitModel"               '-Col F
    SQLx = SQLx & ",UnitType"                '-Col G
    SQLx = SQLx & ",UnitSubType"             '-Col H
    SQLx = SQLx & ",Trans"                   '-Col I
    SQLx = SQLx & ",UnitCost"                '-Col J
    SQLx = SQLx & ",PartsCost"               '-Col K
    SQLx = SQLx & ",OEC"                     '-Col L
    SQLx = SQLx & ",RentLO"                  '-Col M
    SQLx = SQLx & ",IndexType"               '-Col N
    SQLx = SQLx & ",RateTerm"                '-Col O
    SQLx = SQLx & ",RateLO"                  '-Col P
    SQLx = SQLx & ",SwapDate"                '-Col Q
    SQLx = SQLx & ",RVLO"                    '-Col R
    SQLx = SQLx & ",UnitLRF"                 '-Col S
    SQLx = SQLx & ",LTerm"                   '-Col T
    SQLx = SQLx & " FROM vw_PricingRecon_Curr"
    SQLx = SQLx & " WHERE ClientID =" & ClientIDi
    SQLx = SQLx & " AND FAID = '" & FAIDx & "'"
    SQLx = SQLx & " AND GroupID = " & GroupIDi & ";"

    rCntCURR = DCount("FAID", "vw_PricingRecon_CURR", "ClientID = " & ClientIDi & " AND FAID = '" & FAIDx & "' AND GroupID = " & GroupIDi)
    SendReport2Excel SQLz, "PricingRecon", Headerx, True, XLS, rCntLPAF

End Function

Function SynCheckExcel(SchIDx As Long)
    Dim UnitCountx As Integer
    Call MessageUser("Open", "Get Synd Check Data", "This will open Excel...")
    XLS = ""
    XLS = XLS & "|D,BC|G,TC|H,SR|K,$R|L,P2|R,$R|S,$E|T,$R|U,$R|V,$R|W,$E|X,$E|Y,$E"
    XLS = XLS & "|AE,BC|AG,P4|AH,P4|AI,P4|AK,TC|AL,P7|AM,P7|AS,$R|AT,$R|AU,$R|AV,$E|AW,$E|AX,$E"
    XLS = XLS & "|BD,P4|BE,P4|BG,BL|BI,$E|BJ,$E|BK,$E|BL,$E|BM,$E|BN,$E|BO,$E|BP,P7|BT,$E|BU,$E|BV,ZZ|BV,ZZ|A1,DisplayAll|"
    UnitCountx = DCount("AssetID", "ScheduleSyndCheck", "SchID=" & SchIDx)
    SQLz = "Select * From ScheduleSyndCheck Where SchID=" & SchIDx & " Order By SchID, sGrpID;"
    SendReport2Excel SQLz, "SyndCheck", "Schedule Synd Check - " & Format(UnitCountx, "#,#") & " Units " & Format(Date, "mm-dd-yy"), True, XLS, UnitCountx, True
    Call MessageUser("Close")
End Function
Public Function SignOffsPendingReport()

     Call MessageUser("Open", "Gertting Data", "This will take a moment")
    SQLz = "SELECT Client"
    SQLz = SQLz & ",MLNo"
    SQLz = SQLz & ",SCH"
    SQLz = SQLz & ",Grp"
    SQLz = SQLz & ",Assignee"
    SQLz = SQLz & ",FundStatus AS [Fund Status]"
    SQLz = SQLz & ",LastSignoff AS [Last Level]"
    SQLz = SQLz & ",LastSigner"
    SQLz = SQLz & ",LastDate"
    SQLz = SQLz & ",Days"
    SQLz = SQLz & ",NextToSign As [Next To Sign]"
    SQLz = SQLz & " From vw_LastGroupSignOff "
    SQLz = SQLz & " Order by Client, Sch"

    XLS = "|A,DD|d,TC|F,TC|G,TL|J,BC|K,BL|D,W08|F,W07|J,W06|K,W10|"
    Headerx = "Pending Sign-Offs - " & Format(Date, "mm-dd-yy") & " "
    Call MessageUser("Open", "Creating Report ", "...Excel Will Open in a bit...")
        SendReport2Excel SQLz, "Sign-Offs", Headerx, True, XLS, , True
    Call MessageUser("Close")
End Function
Public Function CustDelvReportFinal(Optional Version As Integer)

    Dim Rowi, UnitCnt As Integer
    Dim Endcol As String
        Endcol = "R"
    Call MessageUser("Open", "Opening the Delivery Report", "...This will take moment or two...")
    UnitCnt = DCount("Client", "vw_CustDelvReportFinal")
    Rowi = DCount("Client", "vw_CustDelvReportFinal", "IntDelv=1")
    Rowi = (UnitCnt - Rowi + 4)
    SQLz = "Select  Client"                    '-Col A
    SQLz = SQLz & ", SchGrp As [Sch|Grp]"      '-Col B
    SQLz = SQLz & ", Units As [Cnt]"           '-Col C
    SQLz = SQLz & ", Description"              '-Col D
    SQLz = SQLz & ", Vendor"                   '-Col E
    SQLz = SQLz & ", BuildLast as Build"       '-Col F
    SQLz = SQLz & ", Built"                    '-Col G
    SQLz = SQLz & ", EstDelvDates"             '-Col H
    SQLz = SQLz & ", Delv"                     '-Col I
    SQLz = SQLz & ", MSO"                      '-Col J
    SQLz = SQLz & ", INV As Unit"              '-Col K
    SQLz = SQLz & ", iParts as CP"             '-Col L
    SQLz = SQLz & ", Assignee"                 '-Col M
    SQLz = SQLz & ", PO"                       '-Col N
    SQLz = SQLz & ", IntBank"                  '-Col O
    If Version > 0 Then
        SQLz = SQLz & ",[Last Interim]"        '-Col P
    End If
    SQLz = SQLz & ", COA"                      '-Col Q
    SQLz = SQLz & ", Accept"                   '-Col R
    SQLz = SQLz & ", COADates "                '-Col S
    If Version > 0 Then
        SQLz = SQLz & ", [Sch Out | In]"       '-Col T
        SQLz = SQLz & ", OPSPkg"               '-Col U
        SQLz = SQLz & ", [TMT Sign]"           '-Col V
        SQLz = SQLz & ", [RAL Out | In]"       '-Col W
    End If
    SQLz = SQLz & "  From vw_CustDelvReportFinal "
    SQLz = SQLz & "Order By IntDelv, Client, CDbl(NumInString(Left(SchGrp,3))) ;"




    Headerx = "Customer Delivery Report Summary - " & Format(Date, "mm-dd-yy")
    XLS = "|A,DD3|R,Wrap|B,TX|C,TC|F,TR|G,TR|H,TC|N,TC|"
    XLS = XLS & "C,cLabel:End=E:2:" & TLeftLine & LineSpace(5) & " Unit " & LineSpace(40) & TRightLine & "|"
    XLS = XLS & "D,cLabel:End=D:3:Description|"
    XLS = XLS & "F,cLabel:End=F:3:   Last|"
    XLS = XLS & "F,cLabel:End=G:2:" & TLeftLine & " Build  " & LineSpace(1) & "  Num " & TRightLine & "|"
    XLS = XLS & "H,cLabel:End=H:2:" & TLeftLine & LineSpace(2) & " Est Delivery " & LineSpace(2) & TRightLine & "|"
    XLS = XLS & "H,cLabel:End=H:3:         First ~ Last |"
    XLS = XLS & "K,cLabel:End=L:2:" & TLeftLine & LineSpace(1) & " Invoices " & LineSpace(1) & TRightLine & "|"
    XLS = XLS & "Q,cLabel:End=Q:3:Accepted|"
    XLS = XLS & "R,cLabel:End=R:3:First ~ Last|"
    XLS = XLS & "A,TL|B,TL|D,TL|E,W17|O,TL|"
    If Version > 0 Then
        XLS = XLS & "O,cLabel:End=P:2:" & TLeftLine & LineSpace(2) & "Interim" & LineSpace(5) & TRightLine & "|"
        XLS = XLS & "O,cLabel:End=O:3:Bank|P,cLabel:End=P:3:Last|"
        XLS = XLS & "Q,cLabel:End=Q:3:Sent|"
        XLS = XLS & "Q,cLabel:End=S:2:" & TLeftLine & LineSpace(3) & " COA Sent / Accepted " & LineSpace(4) & TRightLine & "|"
        XLS = XLS & "R,cLabel:End=R:3:Accepted|"
        XLS = XLS & "S,cLabel:End=S:3:First ~ Lst|"
        XLS = XLS & "T,cLabel:End=T:3:SCH|"
        XLS = XLS & "W,cLabel:End=W:3:RAL|"
        XLS = XLS & "T,xlBorder:" & Rowi & ":" & "W|"
        XLS = XLS & "T,TC|U,TC|V,TC|W,TC|"
        Endcol = "W"
    Else
       XLS = XLS & "O,cLabel:End=O:2:" & TLeftLine & LineSpace(1) & "Interim" & LineSpace(2) & TRightLine & "|"
       XLS = XLS & "O,cLabel:End=O:3:Bank|"
       XLS = XLS & "P,cLabel:End=P:3:Sent|"
       XLS = XLS & "P,cLabel:End=R:2:" & TLeftLine & LineSpace(3) & " COA Sent / Accepted " & LineSpace(4) & TRightLine & "|"
    End If
       XLS = XLS & "E,ColFntSize:8|F,ColFntSize:8|G,ColFntSize:8|J,ColFntSize:8|K,ColFntSize:8|L,ColFntSize:8|M,ColFntSize:8|"
       XLS = XLS & "N,ColFntSize:8|O,ColFntSize:8|P,ColFntSize:8|Q,ColFntSize:8|R,ColFntSize:8|G,TC|"
       XLS = XLS & "A,W14|B,W10|C,W10|D,W19|E,W17|E,W16|F,W9|G,W9|H,W16|I,W9|J,W9|K,W9|L,W9|M,W9|N,W5|O,W11|P,W9|Q,W9|R,W11|"
    If Rowi > 0 Then XLS = XLS & "A,InsertRow:" & Rowi & ":|"
    If Rowi > 0 Then XLS = XLS & "A,cLabel:End=" & Endcol & ":" & Rowi & ":" & TLeftLine & LineSpace(70) & " Not Syndicated - Delivered with Interim Funding " & LineSpace(70) & TRightLine & "|"
    XLS = XLS & "A,LL2|"
    XLS = XLS & "A,Banding|"
    If Rowi > 0 Then XLS = XLS & "A,cColor:" & Rowi & ":250.212.52|"
    'XLS = XLS & "E,ColFntSize:8|"
    Call MessageUser("Update", "Updating the Delivery Report", "...Excel will open soon...")
    SendReport2Excel SQLz, "CDReport", Headerx, True, XLS, UnitCnt, 0
    Call MessageUser("Close")
End Function
Public Function SchHealthReport(xPaperSize As String, Optional QuerySource As String)
    Dim Rowi, UnitCnt As Integer
    If QuerySource = "" Then QuerySource = "ScheduleHealthCheckUnFunded"
    UnitCnt = DCount("ClientID", QuerySource)
    'Rowi = DCount("ClientID", "ScheduleHealthCheck")
    Rowi = (UnitCnt - Rowi + 3)
    SQLz = SQLz & "Select"
    SQLz = SQLz & " Lessee"              'Col A
    SQLz = SQLz & ",Schedule"            'Col B
    SQLz = SQLz & ",Delvery"             'Col C
    SQLz = SQLz & ",COA"                 'Col D
    SQLz = SQLz & ",OPSPkg"              'Col E
    SQLz = SQLz & ",SchOutBack"          'Col F
    SQLz = SQLz & ",FinalEconRecv"       'Col G
    SQLz = SQLz & ",TMTSigned"           'Col H
    SQLz = SQLz & ",RALOut"              'Col I
    SQLz = SQLz & ",RALBack"             'Col J
    SQLz = SQLz & ",PermPkgOut"          'Col K
    SQLz = SQLz & ",FATitleTrust"        'Col L
    SQLz = SQLz & ",IntermFundBank as iBank"  'Col M
    SQLz = SQLz & ",IntermFund as iDate"      'Col N
    SQLz = SQLz & ",BLCD"                'Col O
    SQLz = SQLz & ",Assigned"            'Col P
    SQLz = SQLz & ",Assignee"            'Col Q
    SQLz = SQLz & ",FundDate"            'Col R
    SQLz = SQLz & " From " & QuerySource & ";"
    Headerx = "Open Schedule Report - " & Format(Date, "mm-dd-yy")

    XLS = XLS & "|"
    XLS = XLS & "A,mTable:Style=Light1:Name=UnFunded|" 'A,Banding|"
    XLS = XLS & "A,TL|B,Tx|"
    XLS = XLS & "H,VColor:" & Rowi & ":185.211.238|"
    XLS = XLS & "I,VColor:" & Rowi & ":254.214.220|"
    XLS = XLS & "J,VColor:" & Rowi & ":254.214.220|"
    XLS = XLS & "K,VColor:" & Rowi & ":185.211.238|"
    XLS = XLS & "L,VColor:" & Rowi & ":231.230.230|"
    XLS = XLS & "M,VColor:" & Rowi & ":231.230.230|"
    XLS = XLS & "N,VColor:" & Rowi & ":231.230.230|"
    XLS = XLS & "O,VColor:" & Rowi & ":238.203.173|"
    XLS = XLS & "P,VColor:" & Rowi & ":217.225.242|"
    XLS = XLS & "Q,VColor:" & Rowi & ":217.225.242|"
    XLS = XLS & "R,VColor:" & Rowi & ":226.239.218|"
    XLS = XLS & "C,W15|D,W15|H,W15|I,W15|J,W15|N,W15|O,W15|Q,W15|R,W15|"
    XLS = XLS & "I,cLabel:End=J:2:" & TLeftLine & LineSpace(5) & " RAL Letter " & LineSpace(8) & TRightLine & "|"
    XLS = XLS & "M,cLabel:End=N:2:" & TLeftLine & LineSpace(2) & " Interim " & LineSpace(5) & TRightLine & "|"
    XLS = XLS & "A,xlCenter|N,VAutoFit|"
    XLS = XLS & "D,xlBorder:" & Rowi & ":" & "D:Med|"
    XLS = XLS & "H,xlBorder:" & Rowi & ":" & "R:Med|"
    If xPaperSize = "Legal" Then XLS = XLS & "A,LLG|"
    If xPaperSize = "11x17" Then XLS = XLS & "A,LLT|"
    XLS = XLS & "A,TL|H,TL|"
    Call MessageUser("Open", "Making Schedule Health Report", "This will take a moment or two...")
    SendReport2Excel SQLz, "UnFunded Schedules", Headerx, True, XLS, UnitCnt, False
    Call MessageUser("Close")

End Function

Public Function Pause(NumberOfSeconds As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.Description, Erl
    GoTo Exit_GoTo
End Function
