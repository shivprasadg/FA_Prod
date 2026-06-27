Attribute VB_Name = "ExitFunctions"
Option Compare Database
Option Explicit
Public Function MakeRVRealizedReport(Optional ClientGroupIDi As Integer, Optional BankIDi As Integer)
    Dim SQLz As String
    Dim TitleX As String
    Dim XLS As String
    Call MessageUser("Open", "Getting RV Realization", "Just a moment or two...")

    SQLz = "Select ClientGroup" '- Col:A
    SQLz = SQLz & " ,Client" '- Col:B
    SQLz = SQLz & " ,Sch" '- Col:C
    SQLz = SQLz & " ,Group" '- Col:D
    SQLz = SQLz & " ,Assignee" '- Col:E
    SQLz = SQLz & " ,Units" '- Col:F
    SQLz = SQLz & " ,Funded" '- Col:G
    SQLz = SQLz & " ,BLCD" '- Col:H
    SQLz = SQLz & " ,LeaseType" '- Col:I
    SQLz = SQLz & " ,LeaseTerm" '- Col:J
    SQLz = SQLz & " ,LeaseExp" '- Col:K
    SQLz = SQLz & " ,UnitDesc" '- Col:L
    SQLz = SQLz & " ,MPY" '- Col:M
    SQLz = SQLz & " ,AvgODO" '- Col:N
    SQLz = SQLz & " ,Round(AllowedMilesAvg,0)" '- Col:O
    SQLz = SQLz & " ,EOption" '- Col:P
    SQLz = SQLz & " ,RentPassExp" '- Col:Q
    SQLz = SQLz & " ,FirstOffLease" '- Col:R
    SQLz = SQLz & " ,LastOffLease" '- Col:S
    SQLz = SQLz & " ,AgeBLCD" '- Col:T
    SQLz = SQLz & " ,0 as R" '- Col:U
    SQLz = SQLz & " ,VendorCost" '- Col:V
    SQLz = SQLz & " ,SyndRV" '- Col:W
    SQLz = SQLz & " ,SyndRVpct" '- Col:X
    SQLz = SQLz & " ,BookRV" '- Col:Y
    SQLz = SQLz & " ,RmktExp" '- Col:Z
    SQLz = SQLz & " ,SoldPrice" '- Col:AA
    SQLz = SQLz & " ,0 as Y" '- Col:AB
    SQLz = SQLz & " ,0 As Z" '- Col:AC
    SQLz = SQLz & " ,VendorCostSch" '- Col:AD
    SQLz = SQLz & " ,SyndRVSch" '- Col:AE
    SQLz = SQLz & " ,BookRvSch" '- Col:AF
    SQLz = SQLz & " ,RmktExpSch" '- Col:AG
    SQLz = SQLz & " ,SoldPriceSch" '- Col:AH
    SQLz = SQLz & " ,SoldDate" '- Col:AI
    SQLz = SQLz & " ,0 as AF" '- Col:AJ
    SQLz = SQLz & " ,0 as AG" '- Col:AK
    SQLz = SQLz & " ,uNotes" '- Col:AL
    SQLz = SQLz & " FROM vw_RvRealized "
    SQLz = SQLz & " WHERE NOT BLCD IS NULL "
    TitleX = "RV Realized "
    If ClientGroupIDi > 0 Then
        SQLz = SQLz & " AND ClientGroupID= " & ClientGroupIDi & " "
        TitleX = TitleX & " ClientGroup: " & DLookup("ClientGroup", "ClientGroups", "ClientGroupID=" & ClientGroupIDi)
    End If
    If BankIDi > 0 Then
        SQLz = SQLz & " AND AssigneeBankID= " & BankIDi & " "
        TitleX = TitleX & " Bank: " & DLookup("BankShortName", "Banks", "BankID=" & BankIDi)
    End If

    TitleX = TitleX & " : " & Date
    XLS = "|"
    XLS = XLS & "A,TL|A,DD3|A,cLabel:End=A:3:Client Group|A,W08|"
    XLS = XLS & "B,TL|B,DD3|B,cLabel:End=B:3:Client|B,W10|"
    XLS = XLS & "C,TL|C,cLabel:End=C:3:Sch|C,W07|"
    XLS = XLS & "D,TC|D,cLabel:End=D:3:Grp|D,W06|"
    XLS = XLS & "E,TC|E,cLabel:End=E:3:Assignee|E,W07|"
    XLS = XLS & "F,BC|F,cLabel:End=F:3:Units|F,W06|"
    XLS = XLS & "G,TC|G,cLabel:End=G:3:Funded|G,W10|"
    XLS = XLS & "H,TC|H,cLabel:End=H:3:BLCD|H,W10|"
    XLS = XLS & "I,TC|I,cLabel:End=I:3: Lease  Type|I,W07|"
    XLS = XLS & "J,TC|J,cLabel:End=J:3:Primary Lease Term|J,W07|"
    XLS = XLS & "K,BC|K,cLabel:End=K:3:Primary Lease Expire|K,W10|"
    XLS = XLS & "L,TL|L,cLabel:End=L:3:UnitDesc|L,W25|"
    XLS = XLS & "M,Sint|M,cLabel:End=M:3:MPY|M,W09|"
    XLS = XLS & "N,TC|N,Sint|N,cLabel:End=N:3: ODO (avg)|N,W09|"
    XLS = XLS & "O,TC|O,Sint|O,cLabel:End=O:3:Miles Allowed (avg)|O,W09|"
    XLS = XLS & "P,TL|P,cLabel:End=P:3:Exit Option Used|P,W15|"
    XLS = XLS & "Q,$IntE|Q,cLabel:End=Q:3:Rent Pass Expire (Evergreen)|Q,cColor:3:84.130.53|Q,W12|"
    XLS = XLS & "R,TC|R,cLabel:End=R:3:First Off Lease|R,W10|"
    XLS = XLS & "S,TC|S,cLabel:End=S:3:Last Off Lease|S,W10|"
    XLS = XLS & "T,TC|T,cLabel:End=T:3:Age BLCD Mths (avg)|T,cColor:3:0.0.0|T,W09|"
    XLS = XLS & "U,BC|U,cLabel:End=U:3:Age vs Term (avg)|U,cColor:3:117.113.113|U,FxAll=(T4-J4)|U,W07|"
    XLS = XLS & "V,$IntE|V,cLabel:End=V:3:Vendor Cost  (avg)|V,W11|"
    XLS = XLS & "W,$IntE|W,cLabel:End=W:3:Synd RV (avg)|W,cColor:3:191.143.0|W,W12|"
    XLS = XLS & "X,TC|X,P2|X,cLabel:End=X:3:Synd RV (pct)|X,cColor:3:191.143.0|X,FColor:47.117.181|X,W09|"
    XLS = XLS & "Y,$IntE|Y,cLabel:End=Y:3:Book or RV (avg)|Y,cColor:3:191.143.0|Y,W12|"
    XLS = XLS & "Z,$IntE|Z,cLabel:End=Z:3:Rmkt Expense (avg)|Z,cColor:3:191.143.0|Z,W12|"
    XLS = XLS & "AA,$IntE|AA,cLabel:End=AA:3:Sold Price (avg)|AA,cColor:3:191.143.0|AA,W11|"
    XLS = XLS & "AB,$IntE|AB,cLabel:End=AB:3:Net Gain On Sale (avg)|AB,cColor:3:191.143.0|AB,FxAll=(AA4-(W4+Z4))|AB,W12|"
    XLS = XLS & "AC,TC|AC,P2|AC,cLabel:End=AC:3:Sold vs Book or RV (avg)|AC,cColor:3:191.143.0|AC,FColor:47.117.181|AC,FxAll=IFError(AA4/Y4,0.0)|AC,W09|"
    XLS = XLS & "AD,$IntE|AD,cLabel:End=AD:3:Vendor Cost (Sch)|AD,W12|"
    XLS = XLS & "AE,$IntE|AE,cLabel:End=AE:3:Synd RV (SCH)|AE,cColor:3:84.130.53|AE,W12|"
    XLS = XLS & "AF,$IntE|AF,cLabel:End=AF:3:Book or RV (SCH)|AF,cColor:3:84.130.53|AF,W12|"
    XLS = XLS & "AG,$IntE|AG,cLabel:End=AG:3:Rmkt Expense (SCH)|AG,cColor:3:84.130.53|AG,W12|"
    XLS = XLS & "AH,$IntE|AH,cLabel:End=AH:3:Sold Price (SCH)|AH,cColor:3:84.130.53|AH,W12|"
    XLS = XLS & "AI,TC|AI,cLabel:End=AI:3:Sold Date (Max)|AI,cColor:3:0.0.0|AI,W11|"
    XLS = XLS & "AJ,$IntE|AJ,cLabel:End=AJ:3:Gain On Sale (SCH) [Net]|AJ,cColor:3:84.130.53|AJ,FxAll=(AH4-(AG4+AF4))|AJ,W12|"
    XLS = XLS & "AK,TC|AK,P2|AK,cLabel:End=AK:3:Sold VS Book Or RV (SCH)|AK,cColor:3:84.130.53|AK,FColor:47.117.181|AK,FxAll=IfError(AH4/AF4,0)|AK,W07|"
    XLS = XLS & "AL,cLabel:End=AL:3:uNotes|AL,W20|"
    XLS = XLS & "A,Wrap|"

    SendReport2Excel SQLz, "RvRealized", TitleX, True, XLS, 0, False
    Call MessageUser("Close")
End Function
Public Function Make180ExpireReport()
    Dim SQLv As String
    Dim TitleX As String
    Dim XLS As String
    Call MessageUser("Open", "Getting Export", "Just a moment or two...")
    Call MakeExitOptionLocal(True)
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
    SQLv = SQLv & " WHERE AssetID > 0  AND DaysToEOL < 181 Order By DaysToEOL"

    TitleX = "Expiring Units < 181 Days as of: " & Date
    XLS = "|A,BC|C,TC|D,TC|E,BC|F,TC|H,TX|I,BC|K,TC|P,$C|Q,TC|R,BC|S,TC|T,Sint|T,BC|U,TC|Z,TC|"
    XLS = XLS & "AB,$R|AC,$R|AD,$R|AE,SR|AF,TC|AG,TC|AH,TL|AI,TC|AJ,Sint|AJ,TC|"
    XLS = XLS & "AN,TC|AO,TC|AP,TC|AQ,TC|"
      SendReport2Excel SQLv, "EXIT Options", TitleX, True, XLS, 0, False
    Call MessageUser("Close")

End Function
Public Function fSufReport(Optional isTest As Boolean)

    Dim SQLv As String
    Dim RST As DAO.Recordset
    Dim i  As Integer
    Dim NotedUnits As Boolean
    Dim msg1  As String
    Dim m2, m3, m4 As String
    Dim BGColor, WhereX As String
    Dim Filex As String
    Dim Linkx As String
    Dim subjectx As String
    Dim CcEmailX As String
    Dim ContactEmailX As String
    Dim TitleX As String
    Dim XLS As String

    Call MakeExitOptionLocal(True)
    MessageUser "Open", "Making SUF Report", "Just a moment or two..."

    SQLv = ""
    SQLv = SQLv & "SELECT Client"
    SQLv = SQLv & ",Schedule"
    SQLv = SQLv & ",UnitGroup"
    SQLv = SQLv & ",Assignee"
    SQLv = SQLv & ",LeaseType"
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
    SQLv = SQLv & ",SentAge"
    SQLv = SQLv & ",BackAge"
    SQLv = SQLv & ",iDate"
    SQLv = SQLv & ",iPassedDate"
    SQLv = SQLv & ",ActualOffLease"
    SQLv = SQLv & ",LeaseRV"
    SQLv = SQLv & ",BookNow AS [Book (" & Date & ")]"
    SQLv = SQLv & ",BV90Days AS [Book 90d]"
    SQLv = SQLv & ",ODORead"
    SQLv = SQLv & ",ODOSource"
    SQLv = SQLv & ",ODOReadDate"
    SQLv = SQLv & ",Garage"
    SQLv = SQLv & ",ExchDate AS [Can Exch]"
    SQLv = SQLv & ",NewClient"
    SQLv = SQLv & ",NewSchedule"
    SQLv = SQLv & ",NewGroup"
    SQLv = SQLv & ",NewAssetID"
    SQLv = SQLv & ",nInServ"
    SQLv = SQLv & ",ClientID"
    SQLv = SQLv & ",SchID"
    SQLv = SQLv & ",GroupID "
    SQLv = SQLv & ",LeaseOptionID "
    SQLv = SQLv & "FROM SUFReportExport "
    SQLv = SQLv & "WHERE AssetID > 0 "
    'SUF Before Extd
    WhereX = "AND NOT ISNULL(ExtdStart) AND SufSent < ExtdStart AND SUFBack IS NULL Order By Client, SufSent, CurrRentType"
    Set RST = CurrentDb.OpenRecordset(SQLv & WhereX, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        NotedUnits = True
        m2 = "<H3>" & RST.RecordCount & " Units - SUF Sent Before Extension</H3>"
        m2 = m2 & "<table>"
        m2 = m2 & "<TR style=""background-color:#43A047;"" ALIGN=""Center"">"
        m2 = m2 & "<TH>Client</TH>"
        m2 = m2 & "<TH>Sch|Grp</TH>"
        m2 = m2 & "<TH>AssetID</TH>"
        m2 = m2 & "<TH>Unit Desc</TH>"
        m2 = m2 & "<TH>Rent Type</TH>"
        m2 = m2 & "<TH>Unit</br>Status</TH>"
        m2 = m2 & "<TH>Primary</br> Lease Expire</TH>"
        m2 = m2 & "<TH>Days To</br>EOL</TH>"
        m2 = m2 & "<TH>SUF Sent</TH>"
        m2 = m2 & "<TH>SUF</br>Age</TH>"
        m2 = m2 & "<TH style=""background-color:#FF8A65;"" >Extd</br>Start</TH>"
        m2 = m2 & "<TH>Extd</br>End</TH>"
        m2 = m2 & "<TH>SUF Back</TH>"
        m2 = m2 & "<TH>Insp</br>Date</TH>"
        m2 = m2 & "<TH>Insp</br>Passed Date</TH>"
        m2 = m2 & "<TH>Actual</br>Off Lease</TH>"
        m2 = m2 & "</TR>"
        For i = 1 To RST.RecordCount
            BGColor = "Yellow"
            If RST!SentAge > 4 Then BGColor = "#FFC400"
            If RST!SentAge > 7 Then BGColor = "#FF7043"
            If i Mod 2 = 0 Then m2 = m2 & "<TR bgcolor=""WhiteSmoke"">" Else m2 = m2 & "<TR bgcolor=""White"">"
            m2 = m2 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
            m2 = m2 & "<TD ALIGN=""Left"">" & RST!Schedule & "|" & RST!UnitGroup & "</TD>"
            m2 = m2 & "<TD ALIGN=""Left""><b>" & RST!AssetID & "</b></TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!Year & " " & RST!MAKE & " " & RST!MODEL & " " & RST!SubType & " " & RST!Type & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!CurrRentType & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center""><b>" & RST!uStatus & "</b></TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST![Primary Exp] & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center""><b>" & RST!DaysToEOL & "</b></TD>"
            m2 = m2 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!SUFSent & "</TD>"
            m2 = m2 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!SentAge & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!ExtdStart & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!ExtdEnd & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!SUFBack & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!iDate & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!iPassedDate & "</TD>"
            m2 = m2 & "<TD ALIGN=""Center"">" & RST!ActualOffLease & "</TD>"
            m2 = m2 & "</TR>"
            MessageUser "Update", "Making Email", "SUF Before Extd: " & RST.RecordCount & " Units"
            RST.MoveNext
        Next i
        m2 = m2 & "</table><br /><br />"
        RST.close
        Set RST = Nothing
    End If
    'SUF Sent is Stale
    WhereX = "AND SentAge > 2 AND SUFBack IS NULL Order By Client, SufSent, CurrRentType"
    Set RST = CurrentDb.OpenRecordset(SQLv & WhereX, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        NotedUnits = True
        m3 = "<H3>" & RST.RecordCount & " Units - SUF Sent 3+ Months Ago - No SUF Back</H3>"
        m3 = m3 & "<table>"
        m3 = m3 & "<TR style=""background-color:#00ACC1;"" ALIGN=""Center"">"
        m3 = m3 & "<TH>Client</TH>"
        m3 = m3 & "<TH>Sch|Grp</TH>"
        m3 = m3 & "<TH>AssetID</TH>"
        m3 = m3 & "<TH>Unit Desc</TH>"
        m3 = m3 & "<TH>Rent Type</TH>"
        m3 = m3 & "<TH>Unit</br>Status</TH>"
        m3 = m3 & "<TH>Primary</br> Lease Expire</TH>"
        m3 = m3 & "<TH>Days To</br>EOL</TH>"
        m3 = m3 & "<TH>SUF Sent</TH>"
        m3 = m3 & "<TH>Extd</br>Start</TH>"
        m3 = m3 & "<TH>Extd</br>End</TH>"
        m3 = m3 & "<TH style=""background-color:#FF8A65;"">SUF Back</TH>"
        m3 = m3 & "<TH>SUF</br>Age</TH>"
        m3 = m3 & "<TH>Insp</br>Date</TH>"
        m3 = m3 & "<TH>Insp</br>Passed Date</TH>"
        m3 = m3 & "<TH>Actual</br>Off Lease</TH>"
        m3 = m3 & "</TR>"
        For i = 1 To RST.RecordCount
            BGColor = "Yellow"
            If RST!SentAge > 4 Then BGColor = "#FFC400"
            If RST!SentAge > 7 Then BGColor = "#FF7043"
            If i Mod 2 = 0 Then m3 = m3 & "<TR bgcolor=""WhiteSmoke"">" Else m3 = m3 & "<TR bgcolor=""White"">"
            m3 = m3 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
            m3 = m3 & "<TD ALIGN=""Left"">" & RST!Schedule & "|" & RST!UnitGroup & "</TD>"
            m3 = m3 & "<TD ALIGN=""Left""><b>" & RST!AssetID & "</b></TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!Year & " " & RST!MAKE & " " & RST!MODEL & " " & RST!SubType & " " & RST!Type & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!CurrRentType & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center""><b>" & RST!uStatus & "</b></TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST![Primary Exp] & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center""><b>" & RST!DaysToEOL & "</b></TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!SUFSent & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!ExtdStart & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!ExtdEnd & "</TD>"
            m3 = m3 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!SUFBack & "</TD>"
            m3 = m3 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!SentAge & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!iDate & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!iPassedDate & "</TD>"
            m3 = m3 & "<TD ALIGN=""Center"">" & RST!ActualOffLease & "</TD>"
            m3 = m3 & "</TR>"
            MessageUser "Update", "Making Email", "SUF Sent is Stale: " & RST.RecordCount & " Units"
            RST.MoveNext
        Next i
        m3 = m3 & "</table><br /><br />"
        RST.close
        Set RST = Nothing
    End If
    'SUF Back Stale
    WhereX = "AND BackAge > 1 Order By Client, SufBack, CurrRentType"
    Set RST = CurrentDb.OpenRecordset(SQLv & WhereX, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        NotedUnits = True
        m4 = "<H3>" & RST.RecordCount & " Units - SUF Back - Unit Still Active</H3>"
        m4 = m4 & "<table>"
        m4 = m4 & "<TR style=""background-color:#00C853;"" ALIGN=""Center"">"
        m4 = m4 & "<TH>Client</TH>"
        m4 = m4 & "<TH>Sch|Grp</TH>"
        m4 = m4 & "<TH>AssetID</TH>"
        m4 = m4 & "<TH>Unit Desc</TH>"
        m4 = m4 & "<TH>Rent Type</TH>"
        m4 = m4 & "<TH>Unit</br>Status</TH>"
        m4 = m4 & "<TH>Primary</br> Lease Expire</TH>"
        m4 = m4 & "<TH>Days To</br>EOL</TH>"
        m4 = m4 & "<TH>SUF Sent</TH>"
        m4 = m4 & "<TH>Extd</br>Start</TH>"
        m4 = m4 & "<TH>Extd</br>End</TH>"
        m4 = m4 & "<TH>SUF Back</TH>"
        m4 = m4 & "<TH>Insp</br>Date</TH>"
        m4 = m4 & "<TH>Insp</br>Passed Date</TH>"
        m4 = m4 & "<TH>SUF Back</br>Age</TH>"
        m4 = m4 & "<TH style=""background-color:#FF8A65;"">Actual</br>Off Lease</TH>"
        m4 = m4 & "</TR>"
        For i = 1 To RST.RecordCount
            BGColor = "Yellow"
            If RST!BackAge > 4 Then BGColor = "#FFC400"
            If RST!BackAge > 7 Then BGColor = "#FF7043"
            If i Mod 2 = 0 Then m4 = m4 & "<TR bgcolor=""WhiteSmoke"">" Else m4 = m4 & "<TR bgcolor=""White"">"
            m4 = m4 & "<TD ALIGN=""Left"">" & RST!Client & "</TD>"
            m4 = m4 & "<TD ALIGN=""Left"">" & RST!Schedule & "|" & RST!UnitGroup & "</TD>"
            m4 = m4 & "<TD ALIGN=""Left""><b>" & RST!AssetID & "</b></TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!Year & " " & RST!MAKE & " " & RST!MODEL & " " & RST!SubType & " " & RST!Type & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!CurrRentType & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center""><b>" & RST!uStatus & "</b></TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST![Primary Exp] & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center""><b>" & RST!DaysToEOL & "</b></TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!SUFSent & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!ExtdStart & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!ExtdEnd & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!SUFBack & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!iDate & "</TD>"
            m4 = m4 & "<TD ALIGN=""Center"">" & RST!iPassedDate & "</TD>"
            m4 = m4 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!BackAge & "</TD>"
            m4 = m4 & "<TD style=""background-color:" & BGColor & ";"" ALIGN=""Center"">" & RST!ActualOffLease & "</TD>"
            m4 = m4 & "</TR>"
            MessageUser "Update", "Making Email", "SUF Back Stale: " & RST.RecordCount & " Units"
            RST.MoveNext
        Next i
        m4 = m4 & "</table><br /><br />"
        RST.close
        Set RST = Nothing
    End If
    Filex = "\\fa-poseidon\Atlaas\Temp\SUFReport_" & DatePart("M", Date) & "_" & DatePart("D", Date) & "_" & DatePart("YYYY", Date) & ".xlsx"
    Linkx = ""
    Linkx = Filex
    Call KillFiles("\\fa-poseidon\Atlaas\Temp", "SUFReport*.xlsx")
    subjectx = "Notice: IMPORTANT - Possible Invalid SUFs Report"
    If isTest = True Then
        If FAUserName = "ytj" Then
            ContactEmailX = "yjackson@FleetAdvantage.com"
        Else
            ContactEmailX = FAUserName & "@FleetAdvantage.com"
        End If
        CcEmailX = ""
    Else
        ContactEmailX = "Remarketing@FleetAdvantage.com;TMT@FleetAdvantage.com;Fleetservices@fleetadvantage.com"""
        CcEmailX = "it@FleetAdvantage.com;mdeaguiar@fleetadvantage.com;bantonellis@fleetadvantage.com"
    End If

    'make email message
    msg1 = msg1 & "<html><body><H3 style=""color:Navy;font-family:Segoe UI;""><b>EOL ERROR Report - " & Date & "</H3></b></p>"
    msg1 = msg1 & "<p style=""font-family:Calibri;font-size:11pt"">"
    msg1 = msg1 & "Re: " & subjectx & "<br/><br/>"
    msg1 = msg1 & "All,<br/><br/>"
    msg1 = msg1 & "Please review the attached spreadsheet: (<B>" & SplitString(Filex, "Temp\", 1) & "</B>)<br />"
    msg1 = msg1 & "It shows a list of Equipment with SENT SUFs but all are still active.<br /><br />"
    msg1 = msg1 & "All or any required actions are based on your Team`s processes for expiring units.<br />"
    msg1 = msg1 & "If you have any questions or need; i.e.: Mass SUF modifications, please contact the<b> IT Department </b>.<br />"
    If NotedUnits = True Then msg1 = msg1 & "<H3>Units that need attention:</H3>"
    msg1 = msg1 & m2
    msg1 = msg1 & m3
    msg1 = msg1 & m4
    msg1 = msg1 & "If you see any errors or any<i> out of place items</i> within the list of attached units, please contact the IT Department.<br /><br />"
    msg1 = msg1 & "Regards,<br />"
    msg1 = msg1 & "</body></html>"
    TitleX = "Pending SUFs - " & Date
    XLS = ""
    XLS = XLS & "|A,DD|A,WRAP|A,BC"
    XLS = XLS & "|C,TC"
    XLS = XLS & "|D,TC"
    XLS = XLS & "|E,BC"
    XLS = XLS & "|G,TX|G,BC"
    XLS = XLS & "|J,TC"
    XLS = XLS & "|O,$C"
    XLS = XLS & "|P,TC"
    XLS = XLS & "|Q,BC"
    XLS = XLS & "|R,TC"
    XLS = XLS & "|S,Sint|S,BC"
    XLS = XLS & "|T,TL"
    XLS = XLS & "|W,cLabel:End=Z:2:" & TLeftLine & LineSpace(14) & " SUF Details " & LineSpace(14) & TRightLine & ""
    XLS = XLS & "|Y,cLabel:End=Y:3:Sent Age" & vbNewLine & "Mths|Y,BC"
    XLS = XLS & "|Z,cLabel:End=Z:3:Back Age" & vbNewLine & "Mths|Z,BC"
    XLS = XLS & "|AA,cLabel:End=AA:3:INSP" & vbNewLine & "Date|AA,TC"
    XLS = XLS & "|AB,cLabel:End=AB:3:INSP" & vbNewLine & "Pass|AB,TC"
    XLS = XLS & "|AC,$R|AD,$R|AE,$R|AF,$R|AG,SInt|AH,TC|AI,TC"
    XLS = XLS & "|AJ,TL"
    XLS = XLS & "|AK,cLabel:End=AP:2:" & TLeftLine & LineSpace(16) & " EXchangeIT Details (if any) " & LineSpace(25) & TRightLine & ""
    XLS = XLS & "|AL,TC"
    XLS = XLS & "|AO,TC|AP,TC"
    XLS = XLS & "|AQ,HH|AQ,HH|AR,HH|AS,HH|AT,HH"
    XLS = XLS & "|"

    SQLv = SQLv & " Order By Client, SufSent"
    MessageUser "Close"
    MessageUser "Open", "Posting Excel", "still working..."
    SendReport2Excel SQLv, "Sent SUF", TitleX, True, XLS, 0, False, 1, False, , Filex
    MessageUser "Close"
    MessageUser "Open", "Posting Email", "A bit more to go"
    Call SendServerEmail(ContactEmailX, CcEmailX, subjectx, msg1, "ljohnson@fleetadvantage.com", Linkx)
    MessageUser "Close"
End Function
Public Function MakeRVSyndReport(Optional ClientIDi As Long, Optional ClientGroupIDi As Long, Optional ClientX As String, Optional DateStart As Date, Optional DateEnd As Date)
    Dim XLS, Headerx, SQLz As String

    SQLz = "Select Client"
    SQLz = SQLz & ",Schedule"
    SQLz = SQLz & ",GroupNum As [Group]"
    SQLz = SQLz & ",Assignee"
    SQLz = SQLz & ",FundDate AS Funded"
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",UnitVIN"
    SQLz = SQLz & ",UnitStatus As [Status]"
    SQLz = SQLz & ",RentType"
    SQLz = SQLz & ",UnitYr As MYear"
    SQLz = SQLz & ",UnitMake As Make"
    SQLz = SQLz & ",UnitModel As Model"
    SQLz = SQLz & ",UnitTYpe as [Type]"
    SQLz = SQLz & ",UnitSubType as [Sub Type]"
    SQLz = SQLz & ",UnitConfig"
    SQLz = SQLz & ",UnitUOM As UOM"
    SQLz = SQLz & ",Axle"
    SQLz = SQLz & ",AxleConfig"
    SQLz = SQLz & ",OEC"
    SQLz = SQLz & ",LeaseType"
    SQLz = SQLz & ",PLB"
    SQLz = SQLz & ",LeaseTermPrimary As Term"
    SQLz = SQLz & ",RVSales As RVSch"
    SQLz = SQLz & ",RVSynd"
    SQLz = SQLz & ",RVPct"
    SQLz = SQLz & ",RVGuar"
    SQLz = SQLz & ",RVGuarOEM"
    SQLz = SQLz & ",MHAllowed "
    SQLz = SQLz & ",MHCharge "
    SQLz = SQLz & ",AssetID "
    SQLz = SQLz & "From vw_EconoProfileUnit "

   ' If ClientIDi <> 0 Then SQLz = SQLz & " Where ClientID =" & ClientIDi & ""
    If ClientGroupIDi <> 0 Then SQLz = SQLz & " Where ClientGroupID =" & ClientGroupIDi & ""
    XLS = ""
    XLS = XLS & "|A,DD|B,TL|F,TC|G,TL|G,TX|H,BC|I,TC|O,TC|T,BC|R,TC|S,TC"
    XLS = XLS & "|I,cLabel:End=I:3:Rent " & vbNewLine & "Type"
    XLS = XLS & "|O,cLabel:End=O:3:Unit" & vbNewLine & "Config|O,TL"
    XLS = XLS & "|P,TC"
    XLS = XLS & "|Q,TC"
    XLS = XLS & "|R,cLabel:End=R:3:Axle" & vbNewLine & "config|R,TC"
    XLS = XLS & "|S,cLabel:End=S:3:Total OEC" & vbNewLine & "w parts|S,$E"
    XLS = XLS & "|W,cLabel:End=W:3:RV" & vbNewLine & "Sch|W,$R"
    XLS = XLS & "|X,cLabel:End=X:3:RV" & vbNewLine & "Synd|X,$R"
    XLS = XLS & "|Y,cLabel:End=Y:3:RV" & vbNewLine & "%|Y,P2|Y,FColor:47.117.181"
    XLS = XLS & "|Z,cLabel:End=Z:3:FA" & vbNewLine & "Guar|Z,$R|Z,cColor:3:219.168.41"
    XLS = XLS & "|AA,cLabel:End=AA:3:OEM" & vbNewLine & "Guar|AA,$R|AA,cColor:3:183.140.31"
    XLS = XLS & "|A B,cLabel:End=AB:3:Allowed" & vbNewLine & "Miles/Hrs|AB,TC"
    XLS = XLS & "|AC,cLabel:End=AC:3:Overage" & vbNewLine & "Charge|AC,TC"
    XLS = XLS & "|"

    Headerx = "RV Synd Report - All Clients - " & Format(Date, "mm-dd-yy")
    If ClientX <> "" Then Headerx = "RV Synd Report - " & ClientX & " - " & Format(Date, "mm-dd-yy")

    MessageUser "Open", "Excel - RV Syndication Report", "Please Wait this will take a moment..."
    SendReport2Excel SQLz, "RV Synd Report ", Headerx, True, XLS, 0, False
    MessageUser "Close"
End Function
Public Function MakeExitExcelReport(Optional ClientIDi As Long, Optional ClientGroupIDi As Long, Optional ClientX As String, Optional DateStart As Date, Optional DateEnd As Date, Optional MType As Integer, Optional ShortVer As Boolean, Optional Banki As Long)
    Dim XLS, Headerx, SQLz As String

    If ShortVer = True Then
        SQLz = "Select Client" '- Col:A
        SQLz = SQLz & " ,BDEX" '- Col:B
        SQLz = SQLz & " ,Sch" '- Col:C
        SQLz = SQLz & " ,[Group]" '- Col:D
        SQLz = SQLz & " ,Assignee" '- Col:E
        SQLz = SQLz & " ,Funded" '- Col:F
        SQLz = SQLz & " ,FAID" '- Col:G
        SQLz = SQLz & " ,AssetID" '- Col:H
        SQLz = SQLz & " ,UnitNum" '- Col:I
        SQLz = SQLz & " ,UnitVIN" '- Col:J
        SQLz = SQLz & " ,[Status]" '- Col:K
        SQLz = SQLz & " ,InService" '- Col:L
        SQLz = SQLz & " ,Garage" '- Col:M
        SQLz = SQLz & " ,State" '- Col:N
        SQLz = SQLz & " ,UnitYr" '- Col:O
        SQLz = SQLz & " ,UnitMake" '- Col:P
        SQLz = SQLz & " ,UnitModel" '- Col:Q
        SQLz = SQLz & " ,UnitType" '- Col:R
        SQLz = SQLz & " ,UnitSubType" '- Col:S
        SQLz = SQLz & " ,Axle" '- Col:T
        SQLz = SQLz & " ,UnitConfig" '- Col:U
        SQLz = SQLz & " ,Engine" '- Col:V
        SQLz = SQLz & " ,Trans " '- Col:W
        SQLz = SQLz & " ,OEC" '- Col:X
        SQLz = SQLz & " ,LeaseType" '- Col:Y
        SQLz = SQLz & " ,LeaseTerm" '- Col:Z
        SQLz = SQLz & " ,MilesAllowed" '- Col:AA
        SQLz = SQLz & " ,MileCharge" '- Col:AB
        SQLz = SQLz & " ,HoursAllowed" '- Col:AC
        SQLz = SQLz & " ,HourCharge" '- Col:AD
        SQLz = SQLz & " ,BLCD" '- Col:AE
        SQLz = SQLz & " ,LeaseExp" '- Col:AF
        SQLz = SQLz & " ,LRFSch" '- Col:AG
        SQLz = SQLz & " ,RentSch" '- Col:AH
        SQLz = SQLz & " ,LRFSynd" '- Col:AI
        SQLz = SQLz & " ,RentSynd" '- Col:AJ
        SQLz = SQLz & " ,RVSch" '- Col:AK
        SQLz = SQLz & " ,RVSchpct" '- Col:AL
        SQLz = SQLz & " ,RVSynd" '- Col:AM
        SQLz = SQLz & " ,RVpct" '- Col:AN
        SQLz = SQLz & " ,RVGuar" '- Col:AO
        SQLz = SQLz & " ,SyndFeeUnit" '- Col:AP
        SQLz = SQLz & " ,SyndFeeUnitPct" '- Col:AQ
        SQLz = SQLz & " ,SwapType" '- Col:AR
        SQLz = SQLz & " ,SwapYield" '- Col:AS
        SQLz = SQLz & " ,StripRents" '- Col:AT
        SQLz = SQLz & " ,PaidAdvArr" '- Col:AU
        SQLz = SQLz & " ,ImplicitRate" '- Col:AV
        SQLz = SQLz & " ,LRFExtd" '- Col:AW
        SQLz = SQLz & " ,RentExtd" '- Col:AX
        SQLz = SQLz & " ,LRFMTM" '- Col:AY
        SQLz = SQLz & " ,RentMTM" '- Col:AZ
        SQLz = SQLz & " ,UnitSLB" '- Col:BA
        SQLz = SQLz & " ,SUFSent" '- Col:BB
        SQLz = SQLz & " ,SUFBack" '- Col:BC
        SQLz = SQLz & " ,ParkedDate" '- Col:BD
        SQLz = SQLz & " ,ParkedMiles" '- Col:BE
        SQLz = SQLz & " ,HoursParked" '- Col:BF
        SQLz = SQLz & " ,EXchangedBy" '- Col:BG
        SQLz = SQLz & " ,EOptionSelected" '- Col:BH
        SQLz = SQLz & " ,CurrRent" '- Col:BI
        SQLz = SQLz & " ,CurrRentType" '- Col:BJ
        SQLz = SQLz & " ,OffLeaseProj" '- Col:BK
        SQLz = SQLz & " ,AgeBLCD" '- Col:BL
        SQLz = SQLz & " ,ActualOffLease" '- Col:BM
        SQLz = SQLz & " ,ODOSource" '- Col:BN
        SQLz = SQLz & " ,ODORead" '- Col:BO
        SQLz = SQLz & " ,ODODate" '- Col:BP
        SQLz = SQLz & " ,MPY" '- Col:BQ
        SQLz = SQLz & " ,0  as PMiles " ' - Col:BR
        SQLz = SQLz & " ,eType1"       ' - Col:BS
        SQLz = SQLz & " ,eStart1"      ' - Col:BT
        SQLz = SQLz & " ,eTerm1"       ' - Col:BU
        SQLz = SQLz & " ,eEnd1"        ' - Col:BV
        SQLz = SQLz & " ,Rent1"        ' - Col:BW
        SQLz = SQLz & " ,eType2"       ' - Col:BX
        SQLz = SQLz & " ,eStart2"      ' - Col:BY
        SQLz = SQLz & " ,eTerm2"       ' - Col:BZ
        SQLz = SQLz & " ,eEnd2"        ' - Col:CA
        SQLz = SQLz & " ,Rent2"        ' - Col:CB
        SQLz = SQLz & " ,eType3"       ' - Col:CC
        SQLz = SQLz & " ,eStart3"      ' - Col:CD
        SQLz = SQLz & " ,eTerm3"       ' - Col:CE
        SQLz = SQLz & " ,eEnd3"        ' - Col:CF
        SQLz = SQLz & " ,Rent3"        ' - Col:CG
        SQLz = SQLz & " ,Unote"        ' - Col:CH
        SQLz = SQLz & " From vw_ExitReport0 AS E "
        SQLz = SQLz & " Left Join vw_ExtendLeases As X ON E.AssetID = X.UnitID "
        SQLz = SQLz & " WHERE MLOrig='FA' AND  ActualOffLease  IS NULL "

    Else
        SQLz = SQLz & "Select Client"            '- Col:A
        SQLz = SQLz & " ,BDEX"                   '- Col:B
        SQLz = SQLz & " ,MLOrig"                 '- Col:C
        SQLz = SQLz & " ,Sch"                    '- Col:D
        SQLz = SQLz & " ,Group"                  '- Col:E
        SQLz = SQLz & " ,Assignee"               '- Col:F
        SQLz = SQLz & " ,FundProj"               '- Col:G
        SQLz = SQLz & " ,Funded"                 '- Col:H
        SQLz = SQLz & " ,AssetID"                '- Col:I
        SQLz = SQLz & " ,UnitNum"                '- Col:J
        SQLz = SQLz & " ,UnitVIN"                '- Col:K
        SQLz = SQLz & " ,Status"                 '- Col:L
        SQLz = SQLz & " ,Accepted"               '- Col:M
        SQLz = SQLz & " ,EstDelv"                '- Col:N
        SQLz = SQLz & " ,ActualDelivDate"        '- Col:O
        SQLz = SQLz & " ,InService"              '- Col:P
        SQLz = SQLz & " ,PDStart"                '- Col:Q
        SQLz = SQLz & " ,Garage"                 '- Col:R
        SQLz = SQLz & " ,State"                  '- Col:S
        SQLz = SQLz & " ,UnitYr"                 '- Col:T
        SQLz = SQLz & " ,UnitMake"               '- Col:U
        SQLz = SQLz & " ,UnitModel"              '- Col:V
        SQLz = SQLz & " ,UnitType"               '- Col:W
        SQLz = SQLz & " ,UnitSubType"            '- Col:X
        SQLz = SQLz & " ,Axle"                   '- Col:Y
        SQLz = SQLz & " ,UnitConfig"             '- Col:Z
        SQLz = SQLz & " ,Engine"                 '- Col:AA
        SQLz = SQLz & " ,Trans "                 '- Col:AB
        SQLz = SQLz & " ,Unitcost"               '- Col:AC
        SQLz = SQLz & " ,VendorCost"             '- Col:AD
        SQLz = SQLz & " ,ChildPartCost"          '- Col:AE
        SQLz = SQLz & " ,OEC"                    '- Col:AF
        SQLz = SQLz & " ,LeaseType"              '- Col:AG
        SQLz = SQLz & " ,LeaseTerm"              '- Col:AH
        SQLz = SQLz & " ,MilesAllowed"           '- Col:AI
        SQLz = SQLz & " ,MileCharge"             '- Col:AJ
        SQLz = SQLz & " ,HoursAllowed"           '- Col:AK
        SQLz = SQLz & " ,HourCharge"             '- Col:AL
        SQLz = SQLz & " ,BLCD"                   '- Col:AM
        SQLz = SQLz & " ,LRFSch"                 '- Col:AN
        SQLz = SQLz & " ,RentSch"                '- Col:AO
        SQLz = SQLz & " ,LRFSynd"                '- Col:AP
        SQLz = SQLz & " ,RentSynd"               '- Col:AQ
        SQLz = SQLz & " ,ExchangeITMonth"        '- Col:AR
        SQLz = SQLz & " ,ExchangeITDate"         '- Col:AS
        SQLz = SQLz & " ,ExpireNotice"           '- Col:AT
        SQLz = SQLz & " ,LeaseExp"               '- Col:AU
        SQLz = SQLz & " ,RVSch"                  '- Col:AV
        SQLz = SQLz & " ,RVSchpct"               '- Col:AW
        SQLz = SQLz & " ,RVSynd"                 '- Col:AX
        SQLz = SQLz & " ,RVpct"                  '- Col:AY
        SQLz = SQLz & " ,RVGuar"                 '- Col:AZ
        SQLz = SQLz & " ,SyndFeeUnit"            '- Col:BA
        SQLz = SQLz & " ,SyndFeeUnitPct"         '- Col:BB
        SQLz = SQLz & " ,LOIndex"                '- Col:BC
        SQLz = SQLz & " ,LOTerm"                 '- Col:BD
        SQLz = SQLz & " ,LODate"                 '- Col:BE
        SQLz = SQLz & " ,LORate"                 '- Col:BF
        SQLz = SQLz & " ,LOSpread"               '- Col:BG
        SQLz = SQLz & " ,LOYield"                '- Col:BH
        SQLz = SQLz & " ,LOBonus"                '- Col:BI
        SQLz = SQLz & " ,RateAdj"                '- Col:BJ
        SQLz = SQLz & " ,SyndIndex"              '- Col:BK
        SQLz = SQLz & " ,SyndTerm"               '- Col:BL
        SQLz = SQLz & " ,SyndDate"               '- Col:BM
        SQLz = SQLz & " ,SyndRate"               '- Col:BN
        SQLz = SQLz & " ,SyndSpread"             '- Col:BO
        SQLz = SQLz & " ,SyndYield"              '- Col:BP
        SQLz = SQLz & " ,SyndBonus"              '- Col:BQ
        SQLz = SQLz & " ,EXchangedBy"            '- Col:BR
        SQLz = SQLz & " ,EXchangedByUnit"        '- Col:BS
        SQLz = SQLz & " ,EXchangedFor"           '- Col:BT
        SQLz = SQLz & " ,EXchangedForUnit"       '- Col:BU
        SQLz = SQLz & " ,CurrBook"               '- Col:BV
        SQLz = SQLz & " ,SUFSent"                '- Col:BW
        SQLz = SQLz & " ,SUFBack"                '- Col:BX
        SQLz = SQLz & " ,ParkedDate"             '- Col:BY
        SQLz = SQLz & " ,ParkedMiles"            '- Col:BZ
        SQLz = SQLz & " ,HoursParked"            '- Col:CA
        SQLz = SQLz & " ,InspectionHours"        '- Col:CB
        SQLz = SQLz & " ,InspectionDate"         '- Col:CC
        SQLz = SQLz & " ,InspectionPassed"       '- Col:CD
        SQLz = SQLz & " ,EOptionSelected"        '- Col:CE
        SQLz = SQLz & " ,OffLeaseProj"           '- Col:CF
        SQLz = SQLz & " ,ActualOffLease"         '- Col:CG
        SQLz = SQLz & " ,BuyOutDate"             '- Col:CH
        SQLz = SQLz & " ,BuyOutCost"             '- Col:CI
        SQLz = SQLz & " ,TermPenalty"            '- Col:CJ
        SQLz = SQLz & " ,AdjBookValue"           '- Col:CK
        SQLz = SQLz & " ,GainLoss"               '- Col:CL
        SQLz = SQLz & " ,RmktExpense"            '- Col:CM
        SQLz = SQLz & " ,BankShareNet"           '- Col:CN
        SQLz = SQLz & " ,RmktProfit"             '- Col:CO
        SQLz = SQLz & " ,SoldPrice"              '- Col:CP
        SQLz = SQLz & " ,SoldPct"                '- Col:CQ
        SQLz = SQLz & " ,InvoicePaid"            '- Col:CR
        SQLz = SQLz & " ,Buyer"                  '- Col:CS
        SQLz = SQLz & " ,TargetSale"             '- Col:CT
        SQLz = SQLz & " ,TargetSaleDate"         '- Col:CU
        SQLz = SQLz & " ,TermExtd"               '- Col:CV
        SQLz = SQLz & " ,RentExtd"               '- Col:CW
        SQLz = SQLz & " ,LRFExtd"                '- Col:CX
        SQLz = SQLz & " ,RentMTM"                '- Col:CY
        SQLz = SQLz & " ,LRFMTM"                 '- Col:CZ
        SQLz = SQLz & " ,OBCStatus"              '- Col:DA
        SQLz = SQLz & " ,ODOSource"              '- Col:DB
        SQLz = SQLz & " ,AcqMileage"             '- Col:DC
        SQLz = SQLz & " ,ODORead"                '- Col:DD
        SQLz = SQLz & " ,ODODate"                '- Col:DE
        SQLz = SQLz & " ,OBCTotMiles"            '- Col:DF
        SQLz = SQLz & " ,MxOdo"                  '- Col:DG
        SQLz = SQLz & " ,MxOdoDate"              '- Col:DH
        SQLz = SQLz & " ,OBCMPM"                 '- Col:DI
        SQLz = SQLz & " ,OverMiles"              '- Col:DJ
        SQLz = SQLz & " ,OverMilesCost"          '- Col:DK
        SQLz = SQLz & " ,MTM"                    '- Col:DL
        SQLz = SQLz & " ,MPM"                    '- Col:DM
        SQLz = SQLz & " ,MPGOvr"                 '- Col:DN
        SQLz = SQLz & " ,MPGDrv"                 '- Col:DO
        SQLz = SQLz & " ,AgeBLCD"                '- Col:DP
        SQLz = SQLz & " ,CurrRent"               '- Col:DQ
        SQLz = SQLz & " ,CurrRentType"           '- Col:DR
        SQLz = SQLz & " ,RentPassExp"            '- Col:DS
        SQLz = SQLz & " ,TaxRate"                '- Col:DT
        SQLz = SQLz & " ,TaxExempt"              '- Col:DU
        SQLz = SQLz & " ,PaymentsMade"           '- Col:DV
        SQLz = SQLz & " ,BDEXDeal"               '- Col:DW
        SQLz = SQLz & " ,AgeInServ"              '- Col:DX
        SQLz = SQLz & " ,AgeSold"                '- Col:DY
        SQLz = SQLz & " ,SoldDSPN"               '- Col:DZ
        SQLz = SQLz & " ,PODate"                 '- Col:EA
        SQLz = SQLz & " ,POType"                 '- Col:EB
        SQLz = SQLz & " ,UnitSLB"                '- Col:EC
        SQLz = SQLz & " ,StripRents"             '- Col:ED
        SQLz = SQLz & " ,PaidAdvArr"             '- Col:EE
        SQLz = SQLz & " ,FAID"                   '- Col:EF
        SQLz = SQLz & " ,Awarded"                '- Col:EG
        SQLz = SQLz & " ,CPMModel "              '- Col:EH
        SQLz = SQLz & " ,Division "              '- Col:EI
        SQLz = SQLz & " ,Unote "                 '- Col:EJ
        SQLz = SQLz & " ,Extends"                '- Col:EK
        SQLz = SQLz & " ,ClientGroupID "         '- Col:EL
        SQLz = SQLz & " FROM vw_ExitReport0 Where AssetID > 0 "

    End If
    Headerx = "EXIT Report - All Clients - " & Format(Date, "mm-dd-yy")
    If MType = 2 Then SQLz = SQLz & " AND Mlorig='FA'"
    If MType = 3 Then SQLz = SQLz & " AND Mlorig <>'FA'"
    If MType = 4 Then SQLz = SQLz & " AND Mlorig ='FA'"
    If ClientIDi <> 0 Then SQLz = SQLz & " AND ClientID =" & ClientIDi & ""
    If ClientGroupIDi <> 0 Then SQLz = SQLz & " AND ClientGroupID =" & ClientGroupIDi & ""
    If MType = 4 And Nz(DateEnd, "1/1/1900") > "1/1/1900" Then
        SQLz = SQLz & " AND NZ(OffLeaseProj,LeaseExp) >= #" & DateStart & "# and NZ(OffLeaseProj, LeaseExp) <= #" & DateEnd & "# "
        Headerx = Headerx & " FA Only/RMKT Ver."
    End If
    If ClientX <> "" Then Headerx = "EXIT Report - " & ClientX & " - " & Format(Date, "mm-dd-yy")
    If MType = 2 Then Headerx = Headerx & " FA Only"
    If MType = 3 Then Headerx = Headerx & " Tracked Only"
    Headerx = Headerx & " "
    If MType <> 3 Then
        If Banki > 0 Then
            SQLz = SQLz & " AND AssigneeBankId = " & Banki & " "
            Headerx = Headerx & "Assignee:  " & DLookup("BankShortName", "Banks", "BankID=" & Banki) & ""
        End If
    End If
    MessageUser "Open", "Excel - Exit Report", "Please Wait this will take a moment..."

    If ShortVer = True Then
        XLS = ""
        XLS = XLS & "|A,DD3|A,TR|A,mTable:Name=ExitReport"
        XLS = XLS & "|B,BC|B,W05|B,cLabel:End=B:3:BDEX"
        XLS = XLS & "|C,cLabel:End=C:3:Sch"
        XLS = XLS & "|D,W07|D,cLabel:End=D:3:Grp"
        XLS = XLS & "|E,W07|E,cLabel:End=E:3:Assignee"
        XLS = XLS & "|F,TC|F,cLabel:End=F:3:Funded"
        XLS = XLS & "|G,cLabel:End=G:3:FAID"
        XLS = XLS & "|H,BC|H,TX|H,cLabel:End=H:3:AssetID"
        XLS = XLS & "|I,cLabel:End=I:3:Unit #"
        XLS = XLS & "|J,cLabel:End=J:3:VIN"
        XLS = XLS & "|K,BC|K,cLabel:End=K:3:Unit Status|K,cColor:3:0.0.0"
        XLS = XLS & "|L,cLabel:End=L:3:InService"
        XLS = XLS & "|M,cLabel:End=M:3:Garage"
        XLS = XLS & "|N,TC|N,cLabel:End=N:3:State"
        XLS = XLS & "|O,TC|O,cLabel:End=O:3:Model Year"
        XLS = XLS & "|P,cLabel:End=P:3:Make"
        XLS = XLS & "|Q,cLabel:End=Q:3:Model"
        XLS = XLS & "|R,cLabel:End=R:3:UType"
        XLS = XLS & "|S,cLabel:End=S:3:SubType"
        XLS = XLS & "|T,TC|T,cLabel:End=T:3:Axle (type)"
        XLS = XLS & "|U,TC|U,cLabel:End=U:3:Config (UOM)"
        XLS = XLS & "|V,cLabel:End=V:3:Engine"
        XLS = XLS & "|W,cLabel:End=W:3:Trans"
        XLS = XLS & "|X,$E|X,cLabel:End=X:3:Total OEC w parts"
        XLS = XLS & "|Y,TC|Y,cLabel:End=Y:3:Lease Type"
        XLS = XLS & "|Z,TC|Z,cLabel:End=Z:3:Lease Term"
        XLS = XLS & "|AA,SInt|AA,TR|AA,cLabel:End=AA:3:Miles Allowed"
        XLS = XLS & "|AB,$R|AB,cLabel:End=AB:3:Mile Charge"
        XLS = XLS & "|AC,SInt|AC,cLabel:End=AC:3:Hours Allowed"
        XLS = XLS & "|AD,$R|AD,cLabel:End=AD:3:Hour Charge"
        XLS = XLS & "|AE,BC|AE,cLabel:End=AE:3:BLCD"
        XLS = XLS & "|AF,cLabel:End=AF:3:Lease Expire"
        XLS = XLS & "|AG,P3|AG,TC|AG,cLabel:End=AG:3:LRFSch|AG,FColor:47.117.181"
        XLS = XLS & "|AH,$E|AH,cLabel:End=AH:3:RentSch"
        XLS = XLS & "|AI,P2|AI,cLabel:End=AI:3:LRFSynd"
        XLS = XLS & "|AJ,$E|AJ,cLabel:End=AJ:3:RentSynd"
        XLS = XLS & "|AK,$E|AK,cLabel:End=AK:3:RVSch"
        XLS = XLS & "|AL,P2|AL,TC|AL,cLabel:End=AL:3:RVSch %|AL,FColor:47.117.181"
        XLS = XLS & "|AM,$E|AM,cLabel:End=AM:3:RVSynd"
        XLS = XLS & "|AN,P2|AN,TC|AN,cLabel:End=AN:3:RV %|AN,FColor:47.117.181"
        XLS = XLS & "|AO,$E|AO,cLabel:End=AO:3:RV Guar"
        XLS = XLS & "|AP,$E|AP,cLabel:End=AP:3:Synd Fee"
        XLS = XLS & "|AQ,P2|AQ,TC|AQ,cLabel:End=AQ:3:Synd Fee %|AQ,FColor:47.117.181"
        XLS = XLS & "|AR,BC|AR,cLabel:End=AR:3:Swap Type"
        XLS = XLS & "|AS,P3|AS,TC|AS,cLabel:End=AS:3:Swap Yield|AS,cColor:3:128.128.128|AS,FColor:55.86.35"
        XLS = XLS & "|AT,TC|AT,cLabel:End=AT:3:Strip Rents"
        XLS = XLS & "|AU,TC|AU,cLabel:End=AU:3:Paid ADV/ARR"
        XLS = XLS & "|AV,P3|AV,TC|AV,cLabel:End=AV:3:Implicit Rate|AV,FColor:55.86.35"
        XLS = XLS & "|AW,P2|AW,TC|AW,W05|AW,cLabel:End=AW:3:ExtLease LRF|AW,FColor:47.117.181"
        XLS = XLS & "|AX,$E|AX,cLabel:End=AX:3:ExtLease Rent"
        XLS = XLS & "|AY,P2|AY,TC|AY,cLabel:End=AY:3:MTM LRF|AY,FColor:47.117.181"
        XLS = XLS & "|AZ,$E|AZ,cLabel:End=AZ:3:MTM Rent"
        XLS = XLS & "|BA,TC|BA,cLabel:End=BA:3:Unit SLB"
        XLS = XLS & "|BB,cLabel:End=BB:3:SUF Sent"
        XLS = XLS & "|BC,cLabel:End=BC:3:SUF Back"
        XLS = XLS & "|BD,cLabel:End=BD:3:Parked Date"
        XLS = XLS & "|BE,SInt|BE,cLabel:End=BE:3:Parked Miles"
        XLS = XLS & "|BF,SInt|BF,cLabel:End=BF:3:Parked Hours"
        XLS = XLS & "|BG,cLabel:End=BG:3:Exchanged by Sch-Grp|BG,cColor:3:59.112.6"
        XLS = XLS & "|BH,cLabel:End=BH:3:Exit Option|BH,cColor:3:0.0.0"
        XLS = XLS & "|BI,$E|BI,cLabel:End=BI:3:Curr Rent|BI,cColor:3:191.143.0"
        XLS = XLS & "|BJ,TC|BJ,cLabel:End=BJ:3:Rent Type|BJ,cColor:3:191.143.0"
        XLS = XLS & "|BK,cLabel:End=BK:3:Proj Off Lease|BK,cColor:3:59.112.6"
        XLS = XLS & "|BL,TC|BL,cLabel:End=BL:3:Mths In Lease"
        XLS = XLS & "|BM,cLabel:End=BM:3:Actual Off Lease"
        XLS = XLS & "|BN,TC|BN,cLabel:End=BN:3:ODO Source"
        XLS = XLS & "|BO,SInt|BO,cLabel:End=BO:3:ODO Read"
        XLS = XLS & "|BP,cLabel:End=BP:3:ODO Date"
        XLS = XLS & "|BQ,SInt|BQ,cLabel:End=BQ:3:MPY LTD"
        XLS = XLS & "|BQ,cLabel:End=BQ:2:" & Date + 90
        XLS = XLS & "|BR,cLabel:End=BR:3:Proj Miles|BR,FxAll=IF(BO4>0,BO4+((BQ4/12)*(($BQ$2-BP4)/(365/12))),"""")|BR,SInt"
        XLS = XLS & "|BS,TC|BT,TC|BU,TC|BV,TC|BW,$R|BX,TC|BY,TC|BZ,TC|CA,TC|CB,$R|CC,TC|CD,TC|CE,TC|CF,TC|CG,$R"
        XLS = XLS & "|CH,cLabel:End=CH:3:Unote"
        XLS = XLS & "|"

        ' =IF(AND([@Assignee]="FA LLC",$BS$2<[@[Lease Expire]])
        ',FV([@[Swap Yield]]/12
        ',[@['# of months between variable exchange date and lease expiration]],[@RentSch],-[@RVSch]
        ',IF([@[Paid ADV/ARR]]="ADV",1,0)),IF($BS$2<[@[Lease Expire]],FV([@[Implicit Rate]]/12
        ',[@['# of months between variable exchange date and lease expiration]]+[@[Strip Rents]],[@RentSynd],-[@RVSynd]
        ',IF([@[Paid ADV/ARR]]="ADV",1,0)),IF(AND([@[Number of extensions]]=1,[@Assignee]="FA LLC",$BS$2<[@eEnd1]),[@RVSch]-(ROUND(($BS$2-[@eStart1])/(365/12),0)*[@Rent1])
        ',IF(AND([@[Number of extensions]]=1,$BS$2<[@eEnd1]),FV([@[Implicit Rate]]/12,ROUND(($BS$2-[@eStart1])/(365/12),0),[@Rent1],-[@RVSynd]
        ',IF([@[Paid ADV/ARR]]="ADV",1,0)),IF([@[Number of extensions]]=1,FV([@[Implicit Rate]]/12,ROUND(($BS$2-[@eEnd1])/(365/12),0),[@Rent1],-[@[BV 1]],IF([@[Paid ADV/ARR]]="ADV",1,0))
        ',IF(AND([@[Number of extensions]]=2,$BS$2<[@eEnd2]),FV([@[Implicit Rate]]/12,ROUND(($BS$2-[@eStart2])/(365/12),0),[@Rent2],-[@[BV 1]]
        ',IF([@[Paid ADV/ARR]]="ADV",1,0)),IF([@[Number of extensions]]=2,FV([@[Implicit Rate]]/12,ROUND(($BS$2-[@eEnd2])/(365/12),0),[@Rent2],-[@[BV 2]],IF([@[Paid ADV/ARR]]="ADV",1,0)),"")))))))

    Else
        XLS = ""
        XLS = XLS & "|A,DD3|A,TR|A,W15"
        XLS = XLS & "|B,BC|B,W05"
        XLS = XLS & "|C,TC|C,W07"
        XLS = XLS & "|D,TC|D,W07"
        XLS = XLS & "|E,TC|E,W07"
        XLS = XLS & "|F,TC|F,W08"
        XLS = XLS & "|G,TC|G,cLabel:End=G:3:Fund (proj)|G,cColor:3:117.113.113|G,W11"
        XLS = XLS & "|H,TC|H,W11"
        XLS = XLS & "|I,TX|I,BC|I,W08"
        XLS = XLS & "|J,TC|J,W08"
        XLS = XLS & "|K,TL|K,W20"
        XLS = XLS & "|L,cLabel:End=L:3:Unit Status|L,cColor:3:0.0.0|L,cBorder:Med:End=L|L,W06|L,BC"
        XLS = XLS & "|M,TC|M,W11"
        XLS = XLS & "|N,TC|N,cLabel:End=N:3:Delivery Estimate|N,W11"
        XLS = XLS & "|O,TC|O,cLabel:End=O:3:Delivery Actual|O,W11"
        XLS = XLS & "|P,TC|P,W11"
        XLS = XLS & "|Q,TC|Q,W11"
        XLS = XLS & "|R,TL|R,W25"
        XLS = XLS & "|S,TC|S,W05"
        XLS = XLS & "|T,TC|T,cLabel:End=T:3:MYear|T,W06"
        XLS = XLS & "|U,TL|U,cLabel:End=U:3:Make|U,W14"
        XLS = XLS & "|V,TL|V,cLabel:End=V:3:Model|V,W13"
        XLS = XLS & "|W,TL|W,cLabel:End=W:3:UType|W,W11"
        XLS = XLS & "|X,TL|X,cLabel:End=X:3:SubType|X,W12"
        XLS = XLS & "|Y,TC|Y,cLabel:End=Y:3:Axle (type)|Y,W06"
        XLS = XLS & "|Z,TC|Z,cLabel:End=Z:3:Config (UOM)|Z,W08"
        XLS = XLS & "|AA,TC|AA,W15"
        XLS = XLS & "|AB,TC|AB,cLabel:End=AB:3:Trans|AB,W16"
        XLS = XLS & "|AC,$E|AC,cLabel:End=AC:3:Unit Cost|AC,W14"
        XLS = XLS & "|AD,$E|AD,cLabel:End=AD:3:Vendor Cost|AD,W14"
        XLS = XLS & "|AE,$E|AE,cLabel:End=AE:3:Part Cost|AE,W14"
        XLS = XLS & "|AF,$E|AF,cLabel:End=AF:3:Total OEC (w parts)|AF,W14"
        XLS = XLS & "|AG,TC|AG,cLabel:End=AG:3:Lease Type|AG,W09"
        XLS = XLS & "|AH,BC|AH,cLabel:End=AH:3:Lease Term|AH,W06"
        XLS = XLS & "|AI,SR|AI,cLabel:End=AI:3:Miles Allowed|AI,cLabel:End=AL:2:" & TLeftLine & LineSpace(1) & "Allowed M/H Per Year " & LineSpace(10) & TRightLine & "|AI,cColor:2:0.0.128:End=AL|AI,cFontColor:2:255.255.255:End=AL|AI,W10"
        XLS = XLS & "|AJ,$R|AJ,TC|AJ,cLabel:End=AJ:3:Mile Charge|AJ,W06"
        XLS = XLS & "|AK,SR|AK,cLabel:End=AK:3:Hours Allowed|AK,W10"
        XLS = XLS & "|AL,$R|AL,cLabel:End=AL:3:Hour Charge|AL,W06"
        XLS = XLS & "|AM,W11"
        XLS = XLS & "|AN,P3|AN,cLabel:End=AN:3:LRF SCH|AN,W06"
        XLS = XLS & "|AO,$E|AO,cLabel:End=AO:3:Rent SCH|AO,W13"
        XLS = XLS & "|AP,P3|AP,TC|AP,cLabel:End=AP:3:LRF Synd|AP,W06"
        XLS = XLS & "|AQ,$E|AQ,cLabel:End=AQ:3:Rent Synd|AQ,W13"
        XLS = XLS & "|AR,TC|AR,cLabel:End=AR:3:Avail Month|AR,cColor:3:59.112.6:End=AS|AR,cLabel:End=AR:2:" & TLeftLine & LineSpace(1) & " EXch " & LineSpace(6) & TRightLine & "|AR,cColor:2:59.112.6:End=AS|AR,cFontColor:2:255.255.255|AR,W06"
        XLS = XLS & "|AS,TC|AS,cLabel:End=AS:3:Avail Date|AS,cFontColor:2:255.255.255|AS,W11"
        XLS = XLS & "|AT,cLabel:End=AT:3:Expire Notice|AT,W11"
        XLS = XLS & "|AU,cLabel:End=AU:3:Lease Expire|AU,W11"
        XLS = XLS & "|AV,$E|AV,cLabel:End=AV:3:RV SCH|AV,W14"
        XLS = XLS & "|AW,P1|AW,TC|AW,cLabel:End=AW:3:RV SCH %|AW,W06"
        XLS = XLS & "|AX,$E|AX,cLabel:End=AX:3:RV Synd|AX,W14"
        XLS = XLS & "|AY,P1|AY,TC|AY,cLabel:End=AY:3:RV Synd %|AY,W06"
        XLS = XLS & "|AZ,$E|AZ,TC|AZ,cLabel:End=AZ:3:RV Guar|AZ,W10"
        XLS = XLS & "|BA,$E|BA,cLabel:End=BA:3:Synd Fee|BA,W13"
        XLS = XLS & "|BB,P2|BB,TC|BB,cLabel:End=BB:3:Synd Fee %|BB,W06"
        XLS = XLS & "|BC,TC|BC,cLabel:End=BC:3:LO Index|BC,cColor:3:237.125.49:End=BI|BC,cBorder:Med:End=BJ|BC,W07"
        XLS = XLS & "|BD,TC|BD,cLabel:End=BD:3:LO Term YR|BD,W06"
        XLS = XLS & "|BE,TC|BE,cLabel:End=BE:3:LO Date|BE,W11"
        XLS = XLS & "|BF,P3|BF,TC|BF,cLabel:End=BF:3:LO Rate|BF,W09"
        XLS = XLS & "|BG,P3|BG,TC|BG,cLabel:End=BG:3:LO Spread|BG,W09"
        XLS = XLS & "|BH,P3|BH,TC|BH,cLabel:End=BH:3:LO Yield|BH,W09"
        XLS = XLS & "|BI,TC|BI,cLabel:End=BI:3:LO Bonus|BI,W09"
        XLS = XLS & "|BJ,P3|BJ,TC|BJ,cLabel:End=BJ:3:Rate Adj|BJ,cColor:3:128.128.128|BJ,W09"
        XLS = XLS & "|BK,cLabel:End=BK:3:SYND Index|BK,cColor:3:0.0.40:END=BQ|BJ,cBorder:Med:End=BQ|BK,W07|BK,TC"
        XLS = XLS & "|BL,TC|BL,cLabel:End=BL:3:SYND Term YR|BL,W06"
        XLS = XLS & "|BM,TC|BM,cLabel:End=BM:3:SYND Date|BM,W11"
        XLS = XLS & "|BN,P3|BN,TC|BN,cLabel:End=BN:3:SYND Rate|BN,W09"
        XLS = XLS & "|BO,P3|BO,TC|BO,cLabel:End=BO:3: SYND Spread|BO,W09"
        XLS = XLS & "|BP,P3|BP,TC|BP,cLabel:End=BP:3:SYND Yield|BP,W09"
        XLS = XLS & "|BQ,TC|BQ,cLabel:End=BQ:3:SYND Bonus|BQ,W09"
        XLS = XLS & "|BR,TC|BR,cLabel:End=BR:3:Exchanged Sch-Grp|BR,cColor:3:59.112.6|BR,W12"
        XLS = XLS & "|BS,TC|BS,cLabel:End=BS:3:Exch'd Unit|BS,cColor:3:59.112.6|BS,W08"
        XLS = XLS & "|BT,TC|BT,cLabel:End=BT:3:Replaced Sch-Grp|BT,cColor:3:47.117.181|BT,W12"
        XLS = XLS & "|BU,TC|BU,cLabel:End=BU:3:Replaced Unit|BU,cColor:3:47.117.181|BU,W08"
        XLS = XLS & "|BV,$E|BV,cLabel:End=BV:3:Curr Book|BV,cColor:3:0.0.0|BV,cBorder:Med:End=BV|BV,W14"
        XLS = XLS & "|BW,TC|BW,cLabel:End=BW:3:SUF Sent|BW,W11"
        XLS = XLS & "|BX,TC|BX,cLabel:End=BX:3:SUF Back|BX,W11"
        XLS = XLS & "|BY,TC|BY,cLabel:End=BY:3:Parked Date|BY,W11"
        XLS = XLS & "|BZ,Sint|BZ,cLabel:End=BZ:3:Parked Miles|BZ,W09"
        XLS = XLS & "|CA,Sint|CA,cLabel:End=CA:3:Parked Hours|CA,W07"
        XLS = XLS & "|CB,Sint|CB,cLabel:End=CB:3:Insp Hours|CB,W07"
        XLS = XLS & "|CC,TC|CC,cLabel:End=CC:3:Insp Date|CC,cColor:3:0.128.0|CC,W11"
        XLS = XLS & "|CD,TC|CD,cLabel:End=CD:3:Insp Pass|CD,cColor:3:0.100.0|CD,W11"
        XLS = XLS & "|CE,TC|CE,cLabel:End=CE:3:Exit Option|CE,cColor:3:51.51.0|CE,W16"
        XLS = XLS & "|CF,TC|CF,cLabel:End=CF:3:Proj Off Lease|CF,cColor:3:59.112.6|CF,W11"
        XLS = XLS & "|CG,TC|CG,cLabel:End=CG:3:Actual Off Lease|CG,cBorder:Med:End=CG|CG,W11"
        XLS = XLS & "|CH,TC|CH,cLabel:End=CH:3:BuyOut Date|CH,cColor:3:0.0.0|CH,W11"
        XLS = XLS & "|CI,$E|CI,cLabel:End=CI:3:BuyOut Cost|CI,W14"
        XLS = XLS & "|CJ,$E|CJ,cLabel:End=CJ:3:Term Penalty|CJ,W13"
        XLS = XLS & "|CK,$E|CK,cLabel:End=CK:3:FA Book Value|CK,W14"
        XLS = XLS & "|CL,$E|CL,cLabel:End=CL:3:Gain Loss|CL,W12"
        XLS = XLS & "|CM,$E|CM,cLabel:End=CM:3:RMKT Exp|CM,W12"
        XLS = XLS & "|CN,$E|CN,cLabel:End=CN:3:Bank Share|CN,W12"
        XLS = XLS & "|CO,$E|CO,cLabel:End=CO:3:RMKT Profit|CO,W13"
        XLS = XLS & "|CP,$E|CP,cLabel:End=CP:3:Sold Price|CP,cColor:3:55.86.35|CP,cBorder:Med:End=CU|CP,W13"
        XLS = XLS & "|CQ,P2|CQ,TC|CQ,cLabel:End=CQ:3:Sold %|CQ,cColor:3:55.86.35|CQ,W06"
        XLS = XLS & "|CR,TC|CR,cLabel:End=CR:3:Sold Date|CR,cColor:3:55.86.35|CR,W11"
        XLS = XLS & "|CS,BL|CS,cLabel:End=CS:3:RMKT Buyer|CS,cColor:3:84.130.53|CS,W25"
        XLS = XLS & "|CT,$Int|CT,SR|CT,cLabel:End=CT:3:Target Price (FMV)|CT,cColor:3:51.51.0|CT,W09"
        XLS = XLS & "|CU,TC|CU,cLabel:End=CU:3:Target Set Date|CU,cColor:3:51.51.0|CU,W12"
        XLS = XLS & "|CV,TC|CV,cLabel:End=CV:3:ExtLease Term|CV,W07"
        XLS = XLS & "|CW,$E|CW,cLabel:End=CW:3:ExtLease Rent|CW,W12"
        XLS = XLS & "|CX,P3|CX,TC|CX,cLabel:End=CX:3:ExtLease LRF|CX,W07"
        XLS = XLS & "|CY,$R|CY,cLabel:End=CY:3:MTM Rent|CY,W11"
        XLS = XLS & "|CZ,P3|CZ,TC|CZ,cLabel:End=CZ:3:MTM LRF|CZ,W06"
        XLS = XLS & "|DA,TC|DA,cLabel:End=DA:3:OBC Status|DA,W07"
        XLS = XLS & "|DB,TC|DB,cLabel:End=DB:3:ODO Source|DB,W06"
        XLS = XLS & "|DC,Sint|DC,TC|DC,cLabel:End=DC:3:Aquired ODO|DC,cColor:3:123.123.123|DC,W07"
        XLS = XLS & "|DD,Sint|DD,TC|DD,cLabel:End=DD:3:ODO Read|DD,W10"
        XLS = XLS & "|DE,TC|DE,cLabel:End=DE:3:ODO Date|DE,W11"
        XLS = XLS & "|DF,Sint|DF,cLabel:End=DF:3:Cumulative Miles|DF,W10"
        XLS = XLS & "|DG,Sint|DG,cLabel:End=DG:3:Max ODO|DG,W10"
        XLS = XLS & "|DH,TC|DH,cLabel:End=DH:3:Max ODO Date|DH,W11"
        XLS = XLS & "|DI,Sint|DI,cLabel:End=DI:3:MPM Cumulative|DI,W09"
        XLS = XLS & "|DJ,Sint|DJ,cLabel:End=DJ:3:Over Miles Driven|DJ,cColor:3:191.143.0:END=DK|DJ,cBorder:Med:End=DK|DJ,W11"
        XLS = XLS & "|DK,$R|DK,cLabel:End=DK:3:Over Miles Cost|DK,W11"
        XLS = XLS & "|DL,Sint|DL,cLabel:End=DL:3:MPY Last 12|DL,cColor:3:0.0.0:End=DM|DL,W09"
        XLS = XLS & "|DM,Sint|DM,cLabel:End=DM:3:MPM Last 12|DM,W09"
        XLS = XLS & "|DN,TC|DN,cLabel:End=DN:3:MPG OVR|DN,W07"
        XLS = XLS & "|DO,TC|DO,cLabel:End=DO:3:MPG DRV|DO,W07"
        XLS = XLS & "|DP,TC|DP,cLabel:End=DP:3:Mths In Lease|DP,W07"
        XLS = XLS & "|DQ,$E|DQ,cLabel:End=DQ:3:Curr Rent|DQ,cColor:3:191.143.0:End=DR|DQ,W13"
        XLS = XLS & "|DR,TC|DR,cLabel:End=DR:3:Rent Type|DR,W07"
        XLS = XLS & "|DS,$E|DS,cLabel:End=DS:3:Rents Past Expire|DS,W13"
        XLS = XLS & "|DT,P3|DT,TC|DT,cLabel:End=DT:3:Tax Rate|DT,W07"
        XLS = XLS & "|DU,TC|DU,cLabel:End=DU:3:Tax Exempt|DU,W07"
        XLS = XLS & "|DV,cLabel:End=DV:3:Payments Primary / Extened|DV,cColor:3:191.143.0|DV,W25"
        XLS = XLS & "|DW,TC|DW,cLabel:End=DW:3:BDEX Deal|DW,W06"
        XLS = XLS & "|DX,TC|DX,cLabel:End=DX:3:Age InServ|DX,W06"
        XLS = XLS & "|DY,TC|DY,cLabel:End=DY:3:Age Sold|DY,cColor:3:59.112.6|DY,W06"
        XLS = XLS & "|DZ,TC|DZ,cLabel:End=DZ:3:Sold DSPN|DZ,W06"
        XLS = XLS & "|EA,TC|EA,cLabel:End=EA:3:PO Date|EA,W11"
        XLS = XLS & "|EB,TC|EB,cLabel:End=EB:3:PO Type|EB,W05"
        XLS = XLS & "|EC,TC|EC,cLabel:End=EC:3:Unit SLB|EC,W05"
        XLS = XLS & "|ED,TC|ED,cLabel:End=ED:3:Strip Rents|ED,W05"
        XLS = XLS & "|EE,TC|EE,cLabel:End=EE:3:Paid Adv /Arr|EE,W05"
        XLS = XLS & "|EF,TC|EF,W06"
        XLS = XLS & "|EG,TC|EG,W11"
        XLS = XLS & "|EH,TC|EH,cLabel:End=EH:3:MnR Vertical|EH,W14"
        XLS = XLS & "|EI,TC|EI,cLabel:End=EI:3:Division|EI,W12"
        XLS = XLS & "|EJ,TL|EJ,W25"
        XLS = XLS & "|EK,W15"
        XLS = XLS & "|EL,HH"
        XLS = XLS & "|"
    End If
    SendReport2Excel SQLz, "Exit Report ", Headerx, True, XLS, 0, False, , , , "C:\Faas\" & FAUserName & "\Temp\ExitReport.xlsx"
    MessageUser "Close"
End Function

Public Function MakeExitMDAReport(Optional ClientIDi As Long, Optional ClientGroupIDi As Long, Optional ClientX As String, Optional DateStart As Date, Optional DateEnd As Date, Optional MType As Integer, Optional Banki As Long)
    Dim XLS, Headerx, SQLz As String

    SQLz = "Select Client" '- Col:A
    SQLz = SQLz & " ,BDEX" '- Col:B
    SQLz = SQLz & " ,MLOrig" '- Col:C
    SQLz = SQLz & " ,Sch" '- Col:D
    SQLz = SQLz & " ,Group" '- Col:E
    SQLz = SQLz & " ,Assignee" '- Col:F
    SQLz = SQLz & " ,FundProj" '- Col:G
    SQLz = SQLz & " ,LOFQtr" '- Col:H
    SQLz = SQLz & " ,LOFYr" '- Col:I
    SQLz = SQLz & " ,Funded" '- Col:J
    SQLz = SQLz & " ,FundFQtr" '- Col:K
    SQLz = SQLz & " ,FundFYr" '- Col:L
    SQLz = SQLz & " ,AssetID" '- Col:M
    SQLz = SQLz & " ,UnitNum" '- Col:N
    SQLz = SQLz & " ,UnitVIN" '- Col:O
    SQLz = SQLz & " ,Status" '- Col:P
    SQLz = SQLz & " ,Accepted" '- Col:Q
    SQLz = SQLz & " ,EstDelv" '- Col:R
    SQLz = SQLz & " ,ActualDelivDate" '- Col:S
    SQLz = SQLz & " ,InService" '- Col:T
    SQLz = SQLz & " ,PDStart" '- Col:U
    SQLz = SQLz & " ,Garage" '- Col:V
    SQLz = SQLz & " ,State" '- Col:W
    SQLz = SQLz & " ,UnitYr" '- Col:X
    SQLz = SQLz & " ,UnitMake" '- Col:Y
    SQLz = SQLz & " ,UnitModel" '- Col:Z
    SQLz = SQLz & " ,UnitType" '- Col:AA
    SQLz = SQLz & " ,UnitSubType" '- Col:AB
    SQLz = SQLz & " ,Axle" '- Col:AC
    SQLz = SQLz & " ,UnitConfig" '- Col:AD
    SQLz = SQLz & " ,Engine" '- Col:AE
    SQLz = SQLz & " ,Trans " '- Col:AF
    SQLz = SQLz & " ,Unitcost" '- Col:AG
    SQLz = SQLz & " ,VendorCost" '- Col:AH
    SQLz = SQLz & " ,ChildPartCost" '- Col:AI
    SQLz = SQLz & " ,OEC" '- Col:AJ
    SQLz = SQLz & " ,LeaseType" '- Col:AK
    SQLz = SQLz & " ,LeaseTerm" '- Col:AL
    SQLz = SQLz & " ,MilesAllowed" '- Col:AM
    SQLz = SQLz & " ,MileCharge" '- Col:AN
    SQLz = SQLz & " ,HoursAllowed" '- Col:AO
    SQLz = SQLz & " ,HourCharge" '- Col:AP
    SQLz = SQLz & " ,BLCD" '- Col:AQ
    SQLz = SQLz & " ,LRFSch" '- Col:AR
    SQLz = SQLz & " ,RentSch" '- Col:AS
    SQLz = SQLz & " ,LRFSynd" '- Col:AT
    SQLz = SQLz & " ,RentSynd" '- Col:AU
    SQLz = SQLz & " ,ExchangeITMonth" '- Col:AV
    SQLz = SQLz & " ,ExchangeITDate" '- Col:AW
    SQLz = SQLz & " ,ExpireNotice" '- Col:AX
    SQLz = SQLz & " ,LeaseExp" '- Col:AY
    SQLz = SQLz & " ,RVSch" '- Col:AZ
    SQLz = SQLz & " ,RVSchpct" '- Col:BA
    SQLz = SQLz & " ,RVSynd" '- Col:BB
    SQLz = SQLz & " ,RVpct" '- Col:BC
    SQLz = SQLz & " ,RVGuar" '- Col:BD
    SQLz = SQLz & " ,SyndFeeUnit" '- Col:BE
    SQLz = SQLz & " ,SyndFeeUnitPct" '- Col:BF
    SQLz = SQLz & " ,SwapDate" '- Col:BG
    SQLz = SQLz & " ,SwapTerm" '- Col:BH
    SQLz = SQLz & " ,SwapRate" '- Col:BI
    SQLz = SQLz & " ,SwapSpread" '- Col:BJ
    SQLz = SQLz & " ,SwapYield" '- Col:BK
    SQLz = SQLz & " ,SwapBonus" '- Col:BL
    SQLz = SQLz & " ,SwapAdj" '- Col:BM
    SQLz = SQLz & " ,EXchangedBy" '- Col:BN
    SQLz = SQLz & " ,EXchangedByUnit" '- Col:BO
    SQLz = SQLz & " ,EXchangedFor" '- Col:BP
    SQLz = SQLz & " ,EXchangedForUnit" '- Col:BQ
    SQLz = SQLz & " ,CurrBook" '- Col:BR
    SQLz = SQLz & " ,SUFSent" '- Col:BS
    SQLz = SQLz & " ,SUFBack" '- Col:BT
    SQLz = SQLz & " ,ParkedDate" '- Col:BU
    SQLz = SQLz & " ,ParkedMiles" '- Col:BV
    SQLz = SQLz & " ,HoursParked" '- Col:BW
    SQLz = SQLz & " ,InspectionHours" '- Col:BX
    SQLz = SQLz & " ,InspectionDate" '- Col:BY
    SQLz = SQLz & " ,InspectionPassed" '- Col:BZ
    SQLz = SQLz & " ,EOptionSelected" '- Col:CA
    SQLz = SQLz & " ,OffLeaseProj" '- Col:CB
    SQLz = SQLz & " ,ActualOffLease" '- Col:CC
    SQLz = SQLz & " ,BuyOutDate" '- Col:CD
    SQLz = SQLz & " ,BuyOutCost" '- Col:CE
    SQLz = SQLz & " ,TermPenalty" '- Col:CF
    SQLz = SQLz & " ,AdjBookValue" '- Col:CG
    SQLz = SQLz & " ,GainLoss" '- Col:CH
    SQLz = SQLz & " ,RmktExpense" '- Col:CI
    SQLz = SQLz & " ,BankShareNet" '- Col:CJ
    SQLz = SQLz & " ,RmktProfit" '- Col:CK
    SQLz = SQLz & " ,SoldPrice" '- Col:CL
    SQLz = SQLz & " ,SoldPct" '- Col:CM
    SQLz = SQLz & " ,InvoicePaid" '- Col:CN
    SQLz = SQLz & " ,SoldFQTR" '- Col:CO
    SQLz = SQLz & " ,SoldFYr" '- Col:CP
    SQLz = SQLz & " ,Buyer" '- Col:CQ
    SQLz = SQLz & " ,TermExtd" '- Col:CR
    SQLz = SQLz & " ,RentExtd" '- Col:CS
    SQLz = SQLz & " ,LRFExtd" '- Col:CT
    SQLz = SQLz & " ,RentMTM" '- Col:CU
    SQLz = SQLz & " ,LRFMTM" '- Col:CV
    SQLz = SQLz & " ,OBCStatus" '- Col:CW
    SQLz = SQLz & " ,ODOSource" '- Col:CX
    SQLz = SQLz & " ,AcqMileage" '- Col:CY
    SQLz = SQLz & " ,ODORead" '- Col:BZ
    SQLz = SQLz & " ,ODODate" '- Col:DA
    SQLz = SQLz & " ,OBCTotMiles" '- Col:DB
    SQLz = SQLz & " ,MxOdo" '- Col:DC
    SQLz = SQLz & " ,MxOdoDate" '- Col:DD
    SQLz = SQLz & " ,OBCMPM" '- Col:DE
    SQLz = SQLz & " ,MTM" '- Col:DF
    SQLz = SQLz & " ,MPM" '- Col:DG
    SQLz = SQLz & " ,MPGOvr" '- Col:DH
    SQLz = SQLz & " ,MPGDrv" '- Col:DI
    SQLz = SQLz & " ,AgeBLCD" '- Col:DJ
    SQLz = SQLz & " ,CurrRent" '- Col:DK
    SQLz = SQLz & " ,CurrRentType" '- Col:DL
    SQLz = SQLz & " ,RentPassExp" '- Col:DM
    SQLz = SQLz & " ,TaxRate" '- Col:DN
    SQLz = SQLz & " ,TaxExempt" '- Col:DO
    SQLz = SQLz & " ,PaymentsMade" '- Col:DP
    SQLz = SQLz & " ,BDEXDeal" '- Col:DQ
    SQLz = SQLz & " ,AgeInServ" '- Col:DR
    SQLz = SQLz & " ,AgeSold" '- Col:DS
    SQLz = SQLz & " ,SoldDSPN" '- Col:DT
    SQLz = SQLz & " ,PODate" '- Col:DU
    SQLz = SQLz & " ,POType" '- Col:DV
    SQLz = SQLz & " ,UnitSLB" '- Col:DW
    SQLz = SQLz & " ,StripRents" '- Col:DX
    SQLz = SQLz & " ,PaidAdvArr" '- Col:DY
    SQLz = SQLz & " ,FAID" '- Col:DZ
    SQLz = SQLz & " ,Awarded" '- Col:EA
    SQLz = SQLz & " ,CPMModel " '- Col:EB
    SQLz = SQLz & " ,Division " '- Col:EC
    SQLz = SQLz & " ,Unote " '- Col:ED
    SQLz = SQLz & " ,ClientGroupID " '- Col:EE
    SQLz = SQLz & " From vw_ExitReport0 AS E "
    SQLz = SQLz & " WHERE MLOrig='FA' "
    SQLz = SQLz & " AND NOT SCH Like '*RMKT*' "
    SQLz = SQLz & " AND NOT SCH Like '*WARRANTY*' "
    Headerx = "EXIT MDA Report - All Clients - " & Format(Date, "mm-dd-yy")
    MType = 2
'    If MType = 2 Then SQLz = SQLz & " AND Mlorig='FA'"
'    If MType = 3 Then SQLz = SQLz & " AND Mlorig <>'FA'"
'    If MType = 4 Then SQLz = SQLz & " AND Mlorig ='FA'"
    If ClientIDi <> 0 Then SQLz = SQLz & " AND ClientID =" & ClientIDi & ""
    If ClientGroupIDi <> 0 Then SQLz = SQLz & " AND ClientGroupID =" & ClientGroupIDi & ""
    If MType = 4 And Nz(DateEnd, "1/1/1900") > "1/1/1900" Then
        SQLz = SQLz & " AND NZ(OffLeaseProj,LeaseExp) >= #" & DateStart & "# and NZ(OffLeaseProj, LeaseExp) <= #" & DateEnd & "# "
        Headerx = Headerx & " FA Only/RMKT Ver."
    End If
    If ClientX <> "" Then Headerx = "EXIT MDA Report - " & ClientX & " - " & Format(Date, "mm-dd-yy")
    If MType = 2 Then Headerx = Headerx & " FA Only"
   ' If MType = 3 Then Headerx = Headerx & " Tracked Only"
    Headerx = Headerx & " "
    If MType <> 3 Then
        If Banki > 0 Then
            SQLz = SQLz & " AND AssigneeBankId = " & Banki & " "
            Headerx = Headerx & "Assignee:  " & DLookup("BankShortName", "Banks", "BankID=" & Banki) & ""
        End If
    End If
    MessageUser "Open", "Excel - Exit MDA Report", "Please Wait this will take a moment..."

    XLS = "|"
    XLS = XLS & "A,DD3|A,TR|A,cLabel:End=A:3:Client|A,W15|"
    XLS = XLS & "B,BC|B,cLabel:End=B:3:BDEX|B,W05|"
    XLS = XLS & "C,cLabel:End=C:3:MLOrig|C,W07|"
    XLS = XLS & "D,cLabel:End=D:3:Sch|D,W07|"
    XLS = XLS & "E,cLabel:End=E:3:Group|E,W07|"
    XLS = XLS & "F,TC|F,cLabel:End=F:3:Assignee|"
    XLS = XLS & "G,TC|G,cLabel:End=G:3:Fund (proj)|G,cColor:3:117.113.113|G,W12|"
    XLS = XLS & "H,TC|H,cLabel:End=H:3:LO Fiscal QTR|H,cColor:3:117.113.113|H,W07|"
    XLS = XLS & "I,BC|I,cLabel:End=I:3:LO Fiscal YR|I,cColor:3:117.113.113|I,W05|"
    XLS = XLS & "J,TC|J,cLabel:End=J:3:Funded|J,cColor:3:55.86.35|J,W12|"
    XLS = XLS & "K,TC|K,cLabel:End=K:3:Fund Fiscal QTR|K,cColor:3:55.86.35|K,W07|"
    XLS = XLS & "L,BC|L,cLabel:End=L:3:Fund Fiscal YR|L,cColor:3:55.86.35|L,W06|"
    XLS = XLS & "M,TX|M,BC|M,cLabel:End=M:3:AssetID|"
    XLS = XLS & "N,cLabel:End=N:3:UnitNum|"
    XLS = XLS & "O,cLabel:End=O:3:UnitVIN|"
    XLS = XLS & "P,BC|P,cLabel:End=P:3:Unit Status|P,cColor:3:0.0.0|"
    XLS = XLS & "Q,cLabel:End=Q:3:Accepted|"
    XLS = XLS & "R,cLabel:End=R:3:Delivery Estimate|"
    XLS = XLS & "S,cLabel:End=S:3:Delivery Actual|"
    XLS = XLS & "T,cLabel:End=T:3:InService|"
    XLS = XLS & "U,cLabel:End=U:3:PDStart|"
    XLS = XLS & "V,TC|V,cLabel:End=V:3:Garage|V,W05|"
    XLS = XLS & "W,TC|W,cLabel:End=W:3:State|"
    XLS = XLS & "X,TC|X,cLabel:End=X:3:MYear|"
    XLS = XLS & "Y,cLabel:End=Y:3:Make|"
    XLS = XLS & "Z,cLabel:End=Z:3:Model|"
    XLS = XLS & "AA,cLabel:End=AA:3:UType|"
    XLS = XLS & "AB,cLabel:End=AB:3:SubType|"
    XLS = XLS & "AC,TC|AC,cLabel:End=AC:3:Axle (type)|AC,W06|"
    XLS = XLS & "AD,TC|AD,cLabel:End=AD:3:Config (UOM)|"
    XLS = XLS & "AE,0|AE,cLabel:End=AE:3:Engine|"
    XLS = XLS & "AF,TC|AF,cLabel:End=AF:3:Trans|"
    XLS = XLS & "AG,$E|AG,cLabel:End=AG:3:Unit Cost|"
    XLS = XLS & "AH,$E|AH,cLabel:End=AH:3:Vendor Cost|"
    XLS = XLS & "AI,$E|AI,cLabel:End=AI:3:Part Cost|"
    XLS = XLS & "AJ,$E|AJ,cLabel:End=AJ:3:Total OEC (w parts)|"
    XLS = XLS & "AK,cLabel:End=AK:3:Lease Type|"
    XLS = XLS & "AL,BC|AL,cLabel:End=AL:3:Lease Term|AL,W06|"
    XLS = XLS & "AM,SR|AM,cLabel:End=AM:3:Miles Allowed|AM,cLabel:End=AM:2:" & TLeftLine & LineSpace(6) & "|" & "AM,cColor:2:0.0.128|AM,cFontColor:2:255.255.255|AM,W10|"
    XLS = XLS & "AN,$R|AN,TC|AN,cLabel:End=AN:3:Mile Charge|AJ,cLabel:End=AJ:2:Allowed|AJ,cColor:2:0.0.128|AJ,cFontColor:2:255.255.255|AN,W06|"
    XLS = XLS & "AO,SR|AO,cLabel:End=AO:3:Hours Allowed|AK,cLabel:End=AK:2:M/H Per Year|AK,cColor:2:0.0.128|AK,cFontColor:2:255.255.255|AO,W10|"
    XLS = XLS & "AP,$R|AP,cLabel:End=AP:3:Hour Charge|AL,cLabel:End=AL:2:" & LineSpace(5) & TRightLine & "|AL,cColor:2:0.0.128|AL,cFontColor:2:255.255.255|AP,W06|"
    XLS = XLS & "AQ,cLabel:End=AQ:3:BLCD|"
    XLS = XLS & "AR,P3|AR,cLabel:End=AR:3:LRF SCH|AR,cLabel:End=AR:2:" & TLeftLine & LineSpace(1) & " EXch " & "|AR,cColor:2:59.112.6|AR,cFontColor:2:255.255.255|AR,W06|"
    XLS = XLS & "AS,$E|AS,cLabel:End=AS:3:Rent SCH|AS,cLabel:End=AS:2:" & LineSpace(6) & TRightLine & "|AS,cColor:2:59.112.6|AS,cFontColor:2:255.255.255|"
    XLS = XLS & "AT,P3|AT,TC|AT,cLabel:End=AT:3:LRF Synd|AT,FColor:47.117.181|AT,W06|"
    XLS = XLS & "AU,$E|AU,cLabel:End=AU:3:Rent Synd|"
    XLS = XLS & "AV,TC|AV,cLabel:End=AV:3:Avail Month|AV,cColor:3:59.112.6|AV,W06|"
    XLS = XLS & "AW,TC|AW,cLabel:End=AW:3:Avail Date|AW,cColor:3:59.112.6|AW,W10|"
    XLS = XLS & "AX,cLabel:End=AX:3:Expire Notice|"
    XLS = XLS & "AY,cLabel:End=AY:3:Lease Expire|"
    XLS = XLS & "AZ,$E|AZ,cLabel:End=AZ:3:RV SCH|"
    XLS = XLS & "BA,P1|BA,TC|BA,cLabel:End=BA:3:RV SCH %|BA,FColor:47.117.181|"
    XLS = XLS & "BB,$E|BB,cLabel:End=BB:3:RV Synd|"
    XLS = XLS & "BC,P1|BC,TC|BC,cLabel:End=BC:3:RV Synd %|BC,FColor:47.117.181|"
    XLS = XLS & "BD,$E|BD,TC|BD,cLabel:End=BD:3:RV Guar|"
    XLS = XLS & "BE,$E|BE,cLabel:End=BE:3:Synd Fee|"
    XLS = XLS & "BF,P2|BF,TC|BF,cLabel:End=BF:3:Synd Fee %|BF,FColor:47.117.181|"
    XLS = XLS & "BG,TC|BG,cLabel:End=BG:3:Swap Date|BG,cColor:3:128.128.128|"
    XLS = XLS & "BH,TC|BH,cLabel:End=BH:3:Swap Term|BH,cColor:3:128.128.128|"
    XLS = XLS & "BI,P3|BI,TC|BI,cLabel:End=BI:3:Swap Rate|BI,cColor:3:128.128.128|"
    XLS = XLS & "BJ,P3|BJ,TC|BJ,cLabel:End=BJ:3:Swap Spread|BJ,cColor:3:128.128.128|"
    XLS = XLS & "BK,P3|BK,TC|BK,cLabel:End=BK:3:Swap Yield|BK,cColor:3:128.128.128|"
    XLS = XLS & "BL,TC|BL,cLabel:End=BL:3:Swap Bonus|BL,cColor:3:128.128.128|"
    XLS = XLS & "BM,P3|BM,TC|BM,cLabel:End=BM:3:Swap Adj|BM,cColor:3:128.128.128|"
    XLS = XLS & "BN,TC|BN,cLabel:End=BN:3:Exchanged Sch-Grp|BN,cColor:3:59.112.6|"
    XLS = XLS & "BO,TC|BO,cLabel:End=BO:3:Exch'd Unit|BO,cColor:3:59.112.6|BO,W07|"
    XLS = XLS & "BP,TC|BP,cLabel:End=BP:3:Replaced Sch-Grp|BP,cColor:3:47.117.181|"
    XLS = XLS & "BQ,TC|BQ,cLabel:End=BQ:3:Replaced Unit|BQ,cColor:3:47.117.181|"
    XLS = XLS & "BR,$E|BR,cLabel:End=BR:3:Curr Book|BR,cColor:3:0.0.0|"
    XLS = XLS & "BS,TC|BS,cLabel:End=BS:3:SUF Sent|"
    XLS = XLS & "BT,TC|BT,cLabel:End=BT:3:SUF Back|"
    XLS = XLS & "BU,TC|BU,cLabel:End=BU:3:Parked Date|"
    XLS = XLS & "BV,Sint|BV,cLabel:End=BV:3:Parked Miles|"
    XLS = XLS & "BW,Sint|BW,cLabel:End=BW:3:Parked Hours|"
    XLS = XLS & "BX,Sint|BX,cLabel:End=BX:3:Insp Hours|"
    XLS = XLS & "BY,TC|BY,cLabel:End=BY:3:Insp Date|BY,cColor:3:0.128.0|"
    XLS = XLS & "BZ,TC|BZ,cLabel:End=BZ:3:Insp Pass|BZ,cColor:3:0.100.0|"
    XLS = XLS & "CA,TC|CA,cLabel:End=CA:3:Exit Option|CA,cColor:3:0.0.0|"
    XLS = XLS & "CB,TC|CB,cLabel:End=CB:3:Proj Off Lease|CB,cColor:3:59.112.6|"
    XLS = XLS & "CC,TC|CC,cLabel:End=CC:3:Actual Off Lease|"
    XLS = XLS & "CD,TC|CD,cLabel:End=CD:3:BuyOut Date|"
    XLS = XLS & "CE,$E|CE,cLabel:End=CE:3:BuyOut Cost|"
    XLS = XLS & "CF,$E|CF,cLabel:End=CF:3:Term Penalty|"
    XLS = XLS & "CG,$E|CG,cLabel:End=CG:3:FA Book Value|"
    XLS = XLS & "CH,$E|CH,cLabel:End=CH:3:Gain Loss|"
    XLS = XLS & "CI,$E|CI,cLabel:End=CI:3:RMKT Exp|"
    XLS = XLS & "CJ,$E|CJ,cLabel:End=CJ:3:Bank Share|"
    XLS = XLS & "CK,$E|CK,cLabel:End=CK:3:RMKT Profit|"
    XLS = XLS & "CL,$E|CL,cLabel:End=CL:3:Sold Price|"
    XLS = XLS & "CM,P2|CM,TC|CM,cLabel:End=CM:3:Sold %|CM,FColor:47.117.181|"
    XLS = XLS & "CN,TC|CN,cLabel:End=CN:3:Sold Date|CN,W12|"
    XLS = XLS & "CO,TC|CO,cLabel:End=CO:3:Sold Fiscal QTR|CO,W07|"
    XLS = XLS & "CP,BC|CP,cLabel:End=CP:3:Sold Fiscal YR|CP,W06|"
    XLS = XLS & "CQ,BL|CQ,cLabel:End=CQ:3:RMKT Buyer|"
    XLS = XLS & "CR,TC|CR,cLabel:End=CR:3:ExtLease Term|"
    XLS = XLS & "CS,$E|CS,cLabel:End=CS:3:ExtLease Rent|"
    XLS = XLS & "CT,P3|CT,TC|CT,cLabel:End=CT:3:ExtLease LRF|CT,FColor:55.86.35|"
    XLS = XLS & "CU,$R|CU,cLabel:End=CU:3:MTM Rent|"
    XLS = XLS & "CV,P3|CV,TC|CV,cLabel:End=CV:3:MTM LRF|CV,FColor:55.86.35|CV,W07|"
    XLS = XLS & "CW,TC|CW,cLabel:End=CW:3:OBC Status|"
    XLS = XLS & "CX,TC|CX,cLabel:End=CX:3:ODO Source|"
    XLS = XLS & "CY,Sint|CY,TC|CY,cLabel:End=CY:3:Aquired ODO|CY,cColor:3:123.123.123|"
    XLS = XLS & "CZ,Sint|CZ,TC|CZ,cLabel:End=CZ:3:ODO Read|"
    XLS = XLS & "DA,TC|DA,cLabel:End=DA:3:ODO Date|"
    XLS = XLS & "DB,Sint|DB,cLabel:End=DB:3:Cumulative Miles|"
    XLS = XLS & "DC,Sint|DC,cLabel:End=DC:3:Max ODO|"
    XLS = XLS & "DD,TC|DD,cLabel:End=DD:3:Max ODO Date|"
    XLS = XLS & "DE,Sint|DE,cLabel:End=DE:3:MPM Cumulative|"
    XLS = XLS & "DF,Sint|DF,cLabel:End=DF:3:MPY Last 12|DF,cColor:3:0.0.0|"
    XLS = XLS & "DG,Sint|DG,cLabel:End=DG:3:MPM Last 12|DG,cColor:3:0.0.0|"
    XLS = XLS & "DH,TC|DH,cLabel:End=DH:3:MPG OVR|DH,W07|"
    XLS = XLS & "DI,TC|DI,cLabel:End=DI:3:MPG DRV|DI,W07|"
    XLS = XLS & "DJ,TC|DJ,cLabel:End=DJ:3:Mths In Lease|DJ,W07|"
    XLS = XLS & "DK,$E|DK,cLabel:End=DK:3:Curr Rent|DK,cColor:3:191.143.0|"
    XLS = XLS & "DL,TC|DL,cLabel:End=DL:3:Rent Type|DL,cColor:3:191.143.0|DL,W07|"
    XLS = XLS & "DM,$E|DM,cLabel:End=DM:3:Rents Past Expire|DM,W07|"
    XLS = XLS & "DN,P3|DN,TC|DN,cLabel:End=DN:3:Tax Rate|DN,FColor:55.86.35|DN,W07|"
    XLS = XLS & "DO,TC|DO,cLabel:End=DO:3:Tax Exempt|DO,W07|"
    XLS = XLS & "DP,cLabel:End=DP:3:Payments Primary / Extened|DP,cColor:3:191.143.0|"
    XLS = XLS & "DQ,TC|DQ,cLabel:End=DQ:3:BDEX Deal|DQ,W06|"
    XLS = XLS & "DR,TC|DR,cLabel:End=DR:3:Age InServ|DR,W06|"
    XLS = XLS & "DS,TC|DS,cLabel:End=DS:3:Age Sold|DS,cColor:3:59.112.6|DS,W06|"
    XLS = XLS & "DT,TC|DT,cLabel:End=DT:3:Sold DSPN|DT,W06|"
    XLS = XLS & "DU,TC|DU,cLabel:End=DU:3:PO Date|"
    XLS = XLS & "DV,TC|DV,cLabel:End=DV:3:PO Type|DV,W05|"
    XLS = XLS & "DW,TC|DW,cLabel:End=DW:3:Unit SLB|DW,W05|"
    XLS = XLS & "DX,TC|DX,cLabel:End=DX:3:Strip Rents|DX,W05|"
    XLS = XLS & "DY,TC|DY,cLabel:End=DY:3:Paid Adv /Arr|DY,W05|"
    XLS = XLS & "DZ,TC|DZ,cLabel:End=DZ:3:FAID|DZ,W06|"
    XLS = XLS & "EA,TC|EA,cLabel:End=EA:3:Awarded|EA,W06|"
    XLS = XLS & "EB,TC|EB,cLabel:End=EB:3:MnR Vertical|"
    XLS = XLS & "EC,TC|EC,cLabel:End=EC:3:Division |"
    XLS = XLS & "ED,TL|ED,cLabel:End=ED:3:Unote |"
    XLS = XLS & "EE,cLabel:End=EE:3:ClientGroupID|EE,HH|"
    SendReport2Excel SQLz, "Exit Matt Report ", Headerx, True, XLS, 0, False, , , , "C:\Faas\" & FAUserName & "\Temp\ExitMattReport.xlsx"
    MessageUser "Close"
End Function

Public Function MakeExitOptionUnit(UnitIDz As Long, SchIDz As Integer, GroupIDz As Integer, TypeOfExitOption As String)
     If DLookup("MLOrig", "vw_SixKeys", "AssetID=" & UnitIDz & "") <> "FA" Then
        MsgBox "Only FA units can have EXIT Options", , "Not One of Ours"
        Exit Function
    End If
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "AssetID=" & UnitIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "AssetID=" & UnitIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = GroupIDz
    Form_EXITOptions.GroupList_Click
    Form_EXITOptions.RemoteSelectUnit (UnitIDz)
    If TypeOfExitOption = "Extend" Then Form_EXITOptions.ExtendSelected_Click
    If TypeOfExitOption = "Terminate" Then Form_EXITOptions.TerminateSelected_Click
    If TypeOfExitOption = "EXchangeIT" Then Form_EXITOptions.ExchangeSelected_Click
    If TypeOfExitOption = "History" Then Form_EXITOptions.UnitHistory_click
End Function
Public Function MakeExitAMG(Optional ClientGroupIDi As Long, Optional ClientX As String, Optional AssigneeX As String)
    Dim XLS, Headerx, SQLz As String

    SQLz = "Select ClientGroup" '- Col:A
    SQLz = SQLz & " ,Client" '- Col:B
    SQLz = SQLz & " ,Sch" '- Col:C
    SQLz = SQLz & " ,Group" '- Col:D
    SQLz = SQLz & " ,Assignee" '- Col:E
    SQLz = SQLz & " ,LeaseType" '- Col:F
    SQLz = SQLz & " ,SLB" '- Col:G
    SQLz = SQLz & " ,[Unit Status]" '- Col:H
    SQLz = SQLz & " ,Funded" '- Col:I
    SQLz = SQLz & " ,BLCD" '- Col:J
    SQLz = SQLz & " ,[Lease Expiry]" '- Col:K
    SQLz = SQLz & " ,[Actual Off Lease]" '- Col:L
    SQLz = SQLz & " ,[Vendor Cost]" '- Col:M
    SQLz = SQLz & " ,[Part Cost]" '- Col:N
    SQLz = SQLz & " ,[OEC wParts]" '- Col:O
    SQLz = SQLz & " ,[Lease Term]" '- Col:P
    SQLz = SQLz & " ,[Mths In Lease]" '- Col:Q
    SQLz = SQLz & " ,[Adv /Arr]" '- Col:R
    SQLz = SQLz & " ,[Rent Synd]" '- Col:S
    SQLz = SQLz & " ,[Rents Past Expire]" '- Col:T
    SQLz = SQLz & " ,[Payments Primary / Extened]" '- Col:U
    SQLz = SQLz & " ,[RV Synd]" '- Col:V
    SQLz = SQLz & " ,[RV Synd %]" '- Col:W
    SQLz = SQLz & " ,RVGuar" '- Col:X
    SQLz = SQLz & " ,UnitVIN" '- Col:Y
    SQLz = SQLz & " ,MYear" '- Col:Z
    SQLz = SQLz & " ,Make" '- Col:AA
    SQLz = SQLz & " ,Model" '- Col:AB
    SQLz = SQLz & " ,UType" '- Col:AC
    SQLz = SQLz & " ,SubType" '- Col:AD
    SQLz = SQLz & " ,Axle" '- Col:AE
    SQLz = SQLz & " ,[Axle (type)]" '- Col:AF
    SQLz = SQLz & " ,[Config (UOM)]" '- Col:AG
    SQLz = SQLz & " ,Engine" '- Col:AH
    SQLz = SQLz & " ,Trans" '- Col:AI
    SQLz = SQLz & " ,[Miles Allowed]" '- Col:AJ
    SQLz = SQLz & " ,[Mile Charge]" '- Col:AK
    SQLz = SQLz & " ,[Hours Allowed]" '- Col:AL
    SQLz = SQLz & " ,[Hour Charge]" '- Col:AM
    SQLz = SQLz & " ,[Parked Miles]" '- Col:AN
    SQLz = SQLz & " ,[Parked Hours]" '- Col:AO
    SQLz = SQLz & " ,[Equipment returned or replaced]" '- Col:AP
    SQLz = SQLz & " ,[RMKT Exp]" '- Col:AQ
    SQLz = SQLz & " ,[RMKT Sold Price]" '- Col:AR
    SQLz = SQLz & " ,[RMKT Gain on Sale]" '- Col:AS
    SQLz = SQLz & " ,[% Gain on Resale]" '- Col:AT
    SQLz = SQLz & " ,[RMKT Sold Date]" '- Col:AU
    SQLz = SQLz & " ,[Sold DSPN]" '- Col:AV
    SQLz = SQLz & " ,[RMKT Inv Days]" '- Col:AW
    SQLz = SQLz & " ,[MnR Vertical]" '- Col:AX
    SQLz = SQLz & " ,AssetID" '- Col:AY
    SQLz = SQLz & " ,ClientID" '- Col:AZ
    SQLz = SQLz & " ,ClientGroupId" '- Col:BA
    SQLz = SQLz & " FROM vw_ExitAMG WHERE AssetID > 0 "
    Headerx = "EXIT AMG - All Clients - " & Format(Date, "mm-dd-yy")
    If ClientGroupIDi <> 0 Then SQLz = SQLz & " AND ClientGroupID =" & ClientGroupIDi & ""
    If AssigneeX <> "" Then SQLz = SQLz & " AND Assignee ='" & AssigneeX & "'"
    SQLz = SQLz & " ORDER BY Client, Sch "

    If ClientX <> "" Then Headerx = "EXIT AMG - " & ClientX & " - " & Format(Date, "mm-dd-yy")
    MessageUser "Open", "Excel - Exit AMG Report", "Please Wait this will take a moment..."

    XLS = ""
    XLS = XLS & "|A,DD3|A,TR|A,cLabel:End=A:3:ClientGroup|A,TC|A,W10"
    XLS = XLS & "|B,TR|B,cLabel:End=B:3:Client|B,cColor:3:33.92.152|B,TC|B,W11"
    XLS = XLS & "|C,TR|C,cLabel:End=C:3:Sch|C,cColor:3:33.92.152|C,W12"
    XLS = XLS & "|D,TC|D,cLabel:End=D:3:Group|D,cColor:3:33.92.152|D,W10"
    XLS = XLS & "|E,TC|E,cLabel:End=E:3:Assignee|E,cColor:3:33.92.152|E,W10"
    XLS = XLS & "|F,TC|F,cLabel:End=F:3:LeaseType|F,cColor:3:33.92.152|F,W10"
    XLS = XLS & "|G,TC|G,cLabel:End=G:3:SLB|G,cColor:3:33.92.152|G,W05"
    XLS = XLS & "|H,TC|H,cLabel:End=H:3:Unit Status|H,cColor:3:33.92.152|H,W10"
    XLS = XLS & "|I,TC|I,cLabel:End=I:3:Funded|I,cColor:3:33.92.152|I,W10"
    XLS = XLS & "|J,TC|J,cLabel:End=J:3:BLCD|J,cColor:3:33.92.152|J,W10"
    XLS = XLS & "|K,TC|K,cLabel:End=K:3:Lease Expiry|K,cColor:3:255.153.0|K,W10"
    XLS = XLS & "|L,TC|L,cLabel:End=L:3:Actual Off Lease|L,cColor:3:153.102.0|L,W10"
    XLS = XLS & "|M,$R|M,cLabel:End=M:3:Vendor Cost|M,cColor:3:12.118.158|M,W12"
    XLS = XLS & "|N,$R|N,cLabel:End=N:3:Part Cost|N,cColor:3:12.118.158|N,W10"
    XLS = XLS & "|O,$IntE|O,cLabel:End=O:3:OEC wParts|O,cColor:3:12.118.158|O,W14"
    XLS = XLS & "|P,TC|P,cLabel:End=P:3:Lease Term|P,cColor:3:120.33.112|P,W07"
    XLS = XLS & "|Q,TC|Q,cLabel:End=Q:3:Mths In Lease|Q,cColor:3:81.21.74|Q,W07"
    XLS = XLS & "|R,TC|R,cLabel:End=R:3:Adv /Arr|R,cColor:3:15.158.213|R,W07"
    XLS = XLS & "|S,$R|S,cLabel:End=S:3:Rent Synd|S,cColor:3:15.158.213|S,W10"
    XLS = XLS & "|T,$IntE|T,cLabel:End=T:3:Rents Past Expire|T,cColor:3:15.158.213|T,W11"
    XLS = XLS & "|U,TL|U,cLabel:End=U:3:Payments Primary / Extened|U,cColor:3:15.158.213|U,W25"
    XLS = XLS & "|V,$IntE|V,cLabel:End=V:3:RV Synd|V,cColor:3:12.118.158|V,W13"
    XLS = XLS & "|W,P2|W,TC|W,cLabel:End=W:3:RV Synd %|W,cColor:3:12.118.158|W,FColor:47.117.181|W,W08"
    XLS = XLS & "|X,$R|X,cLabel:End=X:3:RVGuar|X,cColor:3:12.118.158|X,W08"
    XLS = XLS & "|Y,TC|Y,cLabel:End=Y:3:UnitVIN|Y,cColor:3:0.0.0|Y,W20"
    XLS = XLS & "|Z,SX|Z,TC|Z,cLabel:End=Z:3:MYear|Z,cColor:3:0.0.0|Z,W07"
    XLS = XLS & "|AA,TL|AA,cLabel:End=AA:3:Make|AA,cColor:3:0.0.0|AA,W15"
    XLS = XLS & "|AB,TL|AB,cLabel:End=AB:3:Model|AB,cColor:3:0.0.0|AB,W20"
    XLS = XLS & "|AC,TL|AC,cLabel:End=AC:3:UType|AC,cColor:3:0.0.0|AC,W15"
    XLS = XLS & "|AD,TL|AD,cLabel:End=AD:3:SubType|AD,cColor:3:0.0.0|AD,W12"
    XLS = XLS & "|AE,0|AE,cLabel:End=AE:3:Axle|AE,cColor:3:0.0.0|AE,W06"
    XLS = XLS & "|AF,TC|AF,cLabel:End=AF:3:Axle (type)|AF,cColor:3:0.0.0|AF,W08"
    XLS = XLS & "|AG,TC|AG,cLabel:End=AG:3:Config (UOM)|AG,cColor:3:0.0.0|AG,W14"
    XLS = XLS & "|AH,TL|AH,cLabel:End=AH:3:Engine|AH,cColor:3:0.0.0|AH,W18"
    XLS = XLS & "|AI,TL|AI,cLabel:End=AI:3:Trans|AI,cColor:3:0.0.0|AI,W14"
    XLS = XLS & "|AJ,SR|AJ,cLabel:End=AJ:3:Miles Allowed|AJ,cColor:3:116.116.116|AJ,W09"
    XLS = XLS & "|AK,$R|AK,cLabel:End=AK:3:Mile Charge|AK,cColor:3:116.116.116|AK,W07"
    XLS = XLS & "|AL,SR|AL,cLabel:End=AL:3:Hours Allowed|AL,cColor:3:116.116.116|AL,W09"
    XLS = XLS & "|AM,$R|AM,cLabel:End=AM:3:Hour Charge|AM,cColor:3:116.116.116|AM,W07"
    XLS = XLS & "|AN,SR|AN,cLabel:End=AN:3:Parked Miles|AN,cColor:3:16.72.97|AN,W08"
    XLS = XLS & "|AO,SR|AO,cLabel:End=AO:3:Parked Hours|AO,cColor:3:16.72.97|AO,W08"
    XLS = XLS & "|AP,TL|AP,cLabel:End=AP:3:Equipment returned or replaced|AP,cColor:3:60.125.34|AP,W20"
    XLS = XLS & "|AQ,$IntE|AQ,cLabel:End=AQ:3:RMKT Exp|AQ,cColor:3:78.167.46|AQ,W12"
    XLS = XLS & "|AR,$IntE|AR,cLabel:End=AR:3:RMKT Sold Price|AR,cColor:3:60.125.34|AR,W12"
    XLS = XLS & "|AS,$IntE|AS,cLabel:End=AS:3:RMKT Gain on Sale|AS,cColor:3:60.125.34|AS,W12"
    XLS = XLS & "|AT,TC|AT,cLabel:End=AT:3:% Gain on Resale|AT,cColor:3:39.83.23|AT,FColor:47.117.181|AT,cFormula:2:=(AR2+T2-AQ2)/V2|AT,%R|AT,W08"
    XLS = XLS & "|AU,TC|AU,cLabel:End=AU:3:RMKT Sold Date|AU,cColor:3:78.167.46|AU,W10"
    XLS = XLS & "|AV,cLabel:End=AV:3:Sold DSPN|AV,cColor:3:78.167.46|AV,W08"
    XLS = XLS & "|AW,BC|AW,cLabel:End=AW:3:RMKT Inv Days|AW,cColor:3:60.125.34|AW,W07"
    XLS = XLS & "|AX,cLabel:End=AX:3:MnR Vertical|AX,cColor:3:116.116.116|AX,W25"
    XLS = XLS & "|AY,HH|AY,cLabel:End=AY:3:AssetID"
    XLS = XLS & "|AZ,HH|AZ,cLabel:End=AZ:3:ClientID"
    XLS = XLS & "|BA,HH|BA,cLabel:End=BA:3:ClientGroupID"
    XLS = XLS & "|"

    SendReport2Excel SQLz, "ExitAMG ", Headerx, True, XLS, 0, False, , , , "C:\Faas\" & FAUserName & "\Temp\ExitAMG.xlsx"
    MessageUser "Close"
End Function
Public Function MakeExitOptionLocal(Optional Override As Boolean, Optional JustUpDate As Boolean)

    If DateDiff("H", Now(), DLookup("WorkDate", "ExitOptionsGroup_Local")) < 8 And Override = False And JustUpDate = False Then Exit Function

    Dim AssetsToUpdate As String
    MessageUser "Open", "Syncing Exit Options", "just a moment..."
'    DoCmd.SetWarnings False
    If JustUpDate = True Then         'Only update the units in the ExitOptions_WorkTable
        Dim RST As DAO.Recordset
        Dim i  As Integer
        Set RST = CurrentDb.OpenRecordset("SELECT UnitID FROM ExitOptions_WorkTable;", dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        Call MessageUser("Update", "Getting AssetIDs", "working...")
        For i = 1 To RST.RecordCount
            AssetsToUpdate = AssetsToUpdate & "," & RST!UnitID
            RST.MoveNext
        Next i
        AssetsToUpdate = Mid(AssetsToUpdate, 2, Len(AssetsToUpdate) - 1)
        RST.close
        Set RST = Nothing
    End If
    Call MessageUser("Update", "Sending Request to Server", "The form will reset when done...")
    ' Calling SP on Server
    Call SetExitOptionsLocal(AssetsToUpdate)

    DoCmd.SetWarnings True
    If IsOpen("GroupForm") Then Forms!GroupForm.Requery
    If IsOpen("EXITOptions") Then
       Forms!ExitOptions.Requery
       Form_EXITOptions.GroupList_Click
    End If
    If IsOpen("ExitOptionsManager") Then
        Forms!ExitOptionsManager.Requery
    End If
    MessageUser "Close"
End Function
