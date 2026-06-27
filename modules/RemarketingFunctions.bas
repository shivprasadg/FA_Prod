Attribute VB_Name = "RemarketingFunctions"
Option Compare Database
 Dim SQLz As String
 Dim SQLx As String
 Dim Committ As String
 Dim XLS As String

Function RMKTInterimStatusReport()

    SQLz = "Select ClientShortName As Client"
    SQLz = SQLz & ",Schedule as SchNo"
    SQLz = SQLz & ",UnitGroup as Grp"
    SQLz = SQLz & ",AssetID"                      ' Col D
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",VIN"
    SQLz = SQLz & ",[Status] "                    ' Col G
    SQLz = SQLz & ",BuyOutDate"
    SQLz = SQLz & ",BuyOutCost"                   ' Col I
    SQLz = SQLz & ",B.BankShortName as  IFBank"
    SQLz = SQLz & ",IFLoanNumber"
    SQLz = SQLz & ",R.IFDate"                     ' Col L
    SQLz = SQLz & ",R.IFAmount"
    SQLz = SQLz & ",R.IFPaidDate"                 ' Col N
    SQLz = SQLz & ",InvoiceSentDate"
    SQLz = SQLz & ",InvoicePaidDate"
    SQLz = SQLz & ",SoldPrice"                    ' Col Q
    SQLz = SQLz & ",BuyDateEffective"
    SQLz = SQLz & " FROM (vw_Sixkeys AS S "
    SQLz = SQLz & " INNER JOIN Rmkt AS R ON S.AssetID = R.UnitRef) "
    SQLz = SQLz & " INNER JOIN Banks AS B ON R.IFBankID = B.BankID"
    SQLz = SQLz & " ORDER BY S.ClientShortName, S.Schedule;"
    MessageUser "Open", "Getting RMKT Interim Status Report ", "This will be a moment..."
    XLS = "|A,DD|C,W15|D,BC|G,BC|F,TC|I,$intE|J,TC|L,TC|M,$intE|N,TC|O,TC|P,TC|Q,$intE|R,TC|"
    Call SendReport2Excel(SQLz, "Interim Status", "RMKT Interim Status Report: " & Date, True, XLS, 0, False)
    MessageUser "Close"

End Function
'_________________________
'XLS = [Col letter],[Format Option][Command]
'Col Letter:        Number/letter of the column
'Format Options:    T = Text | B = Bold | $ = Currency | S = Standard Numbers | P = Percent (P# = Percent and the # of decimials)
'Command:           C = Center | L=Left | R=Right | E = SubTotal Col | T= Total Col Count | X = COUNT | ZZ = Delete column | HH = Hide Column
Function Remarketing_OutStanding_SUF()

    SQLz = "SELECT BankShortName AS Assignee, ClientShortName AS Client "
    SQLz = SQLz & ", MLNo, Schedule, UnitGroup, AssetID, UnitNum, VIN, UnitDescShortVer AS [Unit Desc]"
    SQLz = SQLz & ", [SUF Sent], [SUF Returned], date() - [SUF Sent] As [Days OUT], Status AS UnitStatus, LeaseExp as [Primary Expire], ExtendedEOL as [Extended Expire] "
    SQLz = SQLz & "FROM vw_SixKeys INNER JOIN Rmkt ON vw_SixKeys.UnitID = Rmkt.UnitRef "
    SQLz = SQLz & "WHERE (((Rmkt.[SUF Sent]) Is Not Null) AND ((Rmkt.[SUF Returned]) Is Null)) "
    SQLz = SQLz & "ORDER BY ClientShortName, BankShortName, MLNo, Schedule, UnitNum, VIN, UnitDescShortVer;"
    XLS = "|F,BL|F,TX|L,BC|M,BC|"
    Call SendReport2Excel(SQLz, "Outstanding SUFs", "Outstanding SUFs " & Date, True, XLS, 0, True)
End Function
Function Remarketing_INVReport_Mgr(WHEREString As String)
    Dim SQLx As String
    SQLx = "SELECT"
    SQLx = SQLx & " SourceCompany AS [Account Name]"
    SQLx = SQLx & ",Schedule AS [Lease Sch]"
    SQLx = SQLx & ",Unitgroup AS [Group]"
    SQLx = SQLx & ",UnitNum AS [Customer Unit#]"
    SQLx = SQLx & ",UnitStatus AS [FA Unit Status]"
    SQLx = SQLx & ",UnitYr AS [Year]"
    SQLx = SQLx & ",UnitMake AS Make"
    SQLx = SQLx & ",UnitModel AS Model"
    SQLx = SQLx & ",UnitType AS TYPE"
    SQLx = SQLx & ",[Move From]"
    SQLx = SQLx & ",[Move Start]"
    SQLx = SQLx & ",[Move End]"
    SQLx = SQLx & ",[Move To]"
    SQLx = SQLx & ",ParkedAddress & ' ' & PCity & ' ' & PState AS [Location Addr]"
    SQLx = SQLx & ",ParkDateProjected AS [Projected Park Date]"
    SQLx = SQLx & ",DaysInInventory AS DaysIn"
    SQLx = SQLx & ",InspectedDate AS [Inspection Date]"
    SQLx = SQLx & ",Odometer"
    SQLx = SQLx & ",BuyOut AS [Buy Out Cost]"
    SQLx = SQLx & ",FMV AS FA_FMV"
    SQLx = SQLx & ",PNLFMV AS [PnL (FMV)]"
    SQLx = SQLx & ",TargetSalePrice AS [Target Sale Price]"
    SQLx = SQLx & ",PNLTarget AS [PnL (Target)]"
    SQLx = SQLx & ",GainLoss AS [Term Penalty]"
    SQLx = SQLx & ",BankShareYN AS [Bank Share]"
    SQLx = SQLx & ",IIf(BankShareYN = 'N', NULL,BankPartner) AS [Bank Partner]"
    SQLx = SQLx & ",IIf(BankShareYN = 'N', NULL,BankSharePct) AS [Bank Pct]"
    SQLx = SQLx & ",BankShareAmount AS [Bank $]"
    SQLx = SQLx & ",DaysInAcct AS [Days Since BuyOut]"
    SQLx = SQLx & ",BuyOutDate"
    SQLx = SQLx & ",UnitID"
    SQLx = SQLx & ",ClientID"
    SQLx = SQLx & ",FMV"
    SQLx = SQLx & ",GainLoss"
    SQLx = SQLx & ",TargetSalePrice"
    SQLx = SQLx & ",LocationID"
    SQLx = SQLx & ",PCity"
    SQLx = SQLx & ",PState"
    SQLx = SQLx & ",BuyerID"
    SQLx = SQLx & ",UnitYr"
    SQLx = SQLx & ",UnitMake"
    SQLx = SQLx & ",UnitModel"
    SQLx = SQLx & ",UnitType"
    SQLx = SQLx & ",UnitStatus"
    SQLx = SQLx & ",TitleRecv"
    SQLx = SQLx & ",InvoiceSent"
    SQLx = SQLx & ",[SUF Returned]"
    SQLx = SQLx & ",InvoicePaid"
    SQLx = SQLx & ",Residual"
    SQLx = SQLx & " FROM " & vbNewLine
    SQLx = SQLx & " RemarketingReportAllUnits " & vbNewLine
    SQLx = SQLx & " WHERE " & WHEREString
    XLS = "|A,Wrap|A,DD3|A,BL|E,BC|F,TC|I,TX|P,BC|O,TC|R,SR|S,$E|T,$E|U,$E|V,$E|W,$E|X,$E|Y,TC|Z,TC|AA,P2|AB,$E|AC,TC|AE,BC|AI,$E|AG,$R|AH,$E|AL,TC|AN,TC|AW,$E"
    XLS = XLS & "|A,W12|B,W09|D,W09|E,W08|J,W15|M,W15|N,W20|O,W10|P,W08|Q,W10|V,W14|Y,W05|Z,W05|AA,W07|AC,W09|AK,W16"
    XLS = XLS & "|AF,ZZ|AI,ZZ|AK,ZZ|AO,ZZ|"

    SendReport2Excel SQLx, "Inventory", "Manager's Remarketing Inventory " & Format(UnitCountValue, "#,#") & " Units, as of: " & Format(Date, "mm-dd-yy"), True, XLS, UnitCountValue
End Function

Function Remarketing_INVReport_Sales(SQLString As Variant, UnitCountValue As Integer)

    XLS = "|A,BL|F,BC|G,TC|J,TX|N,W15|O,HY|O,W15|P,SR|Q,$E|R,$E|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,ZZ|S,A0|U,ZZ|U,ZZ|"
    SendReport2Excel SQLString, "Inventory", "Remarketing Inventory " & Format(UnitCountValue, "#,#") & " Units as of: " & Format(Date, "mm-dd-yy"), True, XLS, UnitCountValue
End Function

Function Remarketing_SOLDReport_Mgr(SQLString As Variant, UnitCountValue As Integer)
     XLS = "|D,BR|E,TC|G,BC|H,TX|L,SR|N,TC|P,TC|Q,BC|R,$R|S,$E|T,P2|T,FColor:47.117.181|U,$E|V,$E|W,$E|"
     XLS = XLS & "X,P2|Y,$E|Z,P2|Z,FColor:47.117.181|AA,$E|AB,$E|AD,$E|AF,BC|AG,$E|AJ,TC|AO,$R|"
    SendReport2Excel SQLString, "SoldUnits", "Manager's Sold Units " & Format(UnitCountValue, "#,#") & " Units " & Format(Date, "mm-dd-yy"), True, XLS, UnitCountValue
End Function

Function Remarketing_SOLDReport_Sales(WHEREString As String)
    SQLx = "SELECT SourceCompany AS [Account Name]"                          ' Col A
    SQLx = SQLx & ", Schedule AS [Sch]"                                      ' Col B
    SQLx = SQLx & ", UnitGroup AS [Group]"                                   ' Col C
    SQLx = SQLx & ", UnitID As AssetID"                                      ' Col D
    SQLx = SQLx & ", UnitNum"                                                ' Col E
    SQLx = SQLx & ", UnitVIN AS VIN"                                         ' Col F
    SQLx = SQLx & ", UnitYr AS [Year]"                                       ' Col G
    SQLx = SQLx & ", UnitMake AS Make"                                       ' Col H
    SQLx = SQLx & ", UnitModel AS Model"                                     ' Col I
    SQLx = SQLx & ", UnitType AS Type"                                       ' Col J
    SQLx = SQLx & ", UnitStatus AS [Unit Status]"                            ' Col K
    SQLx = SQLx & ", Odometer"                                               ' Col L
    SQLx = SQLx & ", [SUF Returned]"                                         ' Col M
    SQLx = SQLx & ", PCity AS [Parked City]"                                 ' Col N
    SQLx = SQLx & ", PState AS [Parked State]"                               ' Col O
    SQLx = SQLx & ", SalesPerson AS [Sales Person]"                          ' Col P
    SQLx = SQLx & ", BuyerName"                                              ' Col Q
    SQLx = SQLx & ",[Buyer State]"                                           ' Col R
    SQLx = SQLx & ", BuyersOrder as [Buyers Order]"                          ' Col S
    SQLx = SQLx & ", BuyerType AS Disposition"                               ' Col T
    SQLx = SQLx & ", InvoiceSent As [Invoice Sent]"                          ' Col U
    SQLx = SQLx & ", InvoiceNum AS  [Invoice Num] "                          ' Col V
    SQLx = SQLx & ", InvoicePaid AS  [Invoice Paid] "                        ' Col W
    SQLx = SQLx & ", UnitSalePrice AS [Sale Price]"                          ' Col X
    SQLx = SQLx & ", BuyOutDate AS [BuyOut Date]"                            ' Col Y
    SQLx = SQLx & ", BuyOut AS [BuyOut Cost]"                                ' Col Z
    SQLx = SQLx & ", DaysInAcct AS [Days Since BuyOut]"                      ' Col AA
    SQLx = SQLx & ", 0 AS GainLoss"                                          ' Col AB
    SQLx = SQLx & ", FMV"                                                    ' Col AC
    SQLx = SQLx & ", TargetSalePrice AS [Target Sale Price]"                 ' Col AD
    SQLx = SQLx & ", TitleRecv"                                              ' Col AE
    SQLx = SQLx & ", IFBank"                                                 ' Col AF
    SQLx = SQLx & ", IFDate"                                                 ' Col AG
    SQLx = SQLx & ", IFLoanNumber"                                           ' Col AH
    SQLx = SQLx & ", IFAmount"                                               ' Col AI
    SQLx = SQLx & ", IFPaidDate"                                             ' Col AJ
    SQLx = SQLx & " FROM RemarketingReportAllUnits"
    SQLx = SQLx & " WHERE " & SplitString(WHEREString, " Order By", 0) & " "
    SQLx = SQLx & " ORDER BY SourceCompany, Schedule, UnitGroup;"
    XLS = "|A,mTable:Name=Sales|A,DD3|A,WRAP|A,MedSBL|D,TC|E,BC|G,TC|K,BC|K,TX|L,SintR|O,TC|P,TC|R,TC|S,TC|S,TC|V,TC|X,$E|Z,$E|AA,TC|AB,$R|AC,$R|AD,$R|AF,TC|AH,TC|AI,$E|AJ,TC|"
    XLS = XLS & "D,cColor:3:black|U,cBorder:Med:End=W|U,cColor:3:green:End=X|X,cBorder:Med|AF,cColor:3:crimson:End=AJ|AF,cBorder:Med:End=AJ|AI,cBorder:Med|"
    XLS = XLS & "AB,$E|AB,cColor:3:dark green|AB,cFormula:FxAll=IF(AND([@[Sale Price]]>0,[@[BuyOut Cost]]>0),[@[Sale Price]]-[@[BuyOut Cost]],0)|"
    XLS = XLS & "K,W08|O,W07|P,W07|R,W06|S,W10|T,W09|Y,W11|X,W14|Z,W14|AA,W10|AB,W14|AD,W12|"
    SendReport2Excel SQLx, "SoldUnits", "Sold Units " & Format(UnitCountValue, "#,#") & " Units " & Format(Date, "mm-dd-yy"), True, XLS, 0
End Function

Function Remarketing_AdminReport(SQLString As Variant, UnitCountValue As Integer)
    XLS = "|A,TL|C,TC|D,TX|K,BC|G,BC|O,BC|Y,BC|Z,$E|AA,$E|AB,$E|AC,$E|AD,$E|AE,$E|AJ,TC|"
    SendReport2Excel SQLString, "Remarketing", "Admin Report " & Format(UnitCountValue, "#,#") & " Units " & Format(Date, "mm-dd-yy"), True, XLS, 0
End Function

Function ScoreRemarketingUnits(SQLStatement As String)
    Set RST = CurrentDb.OpenRecordset(SQLStatement)
    RST.MoveFirst
    RST.MoveLast
    UnitCountx = RST.RecordCount
    RST.MoveFirst
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From RemarketingSales_WorkTable;")
    For i = 1 To UnitCountx
        DoCmd.RunSQL ("INSERT INTO RemarketingSales_WorkTable ( UnitID ) SELECT " & RST!UnitID & ";")
        RST.MoveNext
    Next i
    DoCmd.SetWarnings True
    Set RST = Nothing
End Function
Function RMKT_BankInventoryReport(iBankId As Long, Optional sBankName As String)

    SQLz = "Select SourceCompany AS [Source Client]"                        '- Col:A
    SQLz = SQLz & " ,BankPartner AS Bank"                                   '- Col:B
    SQLz = SQLz & " ,UnitVin AS VIN"                                        '- Col:C
    SQLz = SQLz & " ,UnitStatus AS [Unit Status]"                           '- Col:D
    SQLz = SQLz & " ,UnitDescShortVer AS [Unit Desc]"                       '- Col:E
    SQLz = SQLz & " ,PState"                                                '- Col:F
    SQLz = SQLz & " ,LeaseExp AS [Lease Expire]"                            '- Col:G
    SQLz = SQLz & " ,OffLeaseDate AS [Actual OffLease]"                     '- Col:H
    SQLz = SQLz & " ,InvoicePaid AS [Sold Date]"                            '- Col:I
    SQLz = SQLz & " ,LeaseTermPrimary AS [Lease Term]"                      '- Col:J
    SQLz = SQLz & " ,AgeAtSale AS [Age At Sale] "                           '- Col:K
    SQLz = SQLz & " ,Residual AS RV"                                        '- Col:L
    SQLz = SQLz & " ,CurrBook AS [Book at Sale]"                            '- Col:M
    SQLz = SQLz & " ,BuyOut"                                                '- Col:N
    SQLz = SQLz & " ,TotalExpense AS [RMKT Exp]"                            '- Col:O
    SQLz = SQLz & " ,UnitSalePrice AS [Sold Price]"                         '- Col:P
    SQLz = SQLz & " ,TotalProfit AS [Gain Loss]"                            '- Col:Q
    SQLz = SQLz & " ,BankShareYN AS [Bank Share Y/N]"                       '- Col:R
    SQLz = SQLz & " ,iif(BankShare < 1,Null,BankShare) AS [Bank Share]"     '- Col:S
    SQLz = SQLz & " ,RentPassExp AS [Evergreen Rent]"                       '- Col:T
    SQLz = SQLz & " ,Odometer AS ODO"                                       '- Col:U
    SQLz = SQLz & " ,Milesallowed AS [Miles Allowed]"                       '- Col:V
    SQLz = SQLz & " ,iif(OverMiles<0,Null,OverMiles) AS [Over Miles]"       '- Col:W
    SQLz = SQLz & " ,iif(OverMilesCost=0,Null,OverMilesCost) AS [Over Miles Cost]"    '- Col:X
    SQLz = SQLz & " ,Disposition AS Disp"                                   '- Col:Y
    SQLz = SQLz & " ,[Sch|Grp] AS [Lease Sch|Grp]"                          '- Col:Z
    SQLz = SQLz & " ,UnitId AS AssetID"                                     '- Col:AA
    SQLz = SQLz & " ,Unitnum"                                               '- Col:AB
    SQLz = SQLz & " ,UnitSubType AS SubType"                                '- Col:AC
    SQLz = SQLz & " ,ParkedLocation As [Parked Location]"                   '- Col:AD
    SQLz = SQLz & " ,EOption As [Exit Options]"                             '- Col:AE
    SQLz = SQLz & " FROM vw_RemarketingReportAllUnits "
    SQLz = SQLz & " WHERE UnitStatus IN ('S', 'D') AND InvoicePaid > Date() - 766 "
    If iBankId > 0 Then SQLz = SQLz & "AND BankID = " & iBankId & " "
    SQLz = SQLz & " ORDER BY SourceCompany, BankPartner;"

    XLS = "|A,mTable:Style=Medium4:Name=SoldUnits|A,DD3|A,W12|"
    XLS = XLS & "B,TL|B,W10|"
    XLS = XLS & "C,TC|C,TX|C,W19|"
    XLS = XLS & "D,BC|D,cColor:3:0.0.0|D,cBorder:Med|D,W06|"
    XLS = XLS & "E,TL|E,W20|"
    XLS = XLS & "F,TC|F,W06|"
    XLS = XLS & "G,TC|G,W10|"
    XLS = XLS & "H,TC|H,W10|"
    XLS = XLS & "I,TC|I,W10|"
    XLS = XLS & "J,TC|J,cBorder:Med:End=K|J,cColor:3:0.0.0|J,W06|"
    XLS = XLS & "K,BC|K,cColor:3:0.0.0|K,W06|"
    XLS = XLS & "L,$IntE|L,cColor:3:55.86.35|L,W12|"
    XLS = XLS & "M,$IntE|M,cColor:3:55.86.35|M,W12|"
    XLS = XLS & "N,$IntE|N,cColor:3:55.86.35|N,W12|"
    XLS = XLS & "O,$IntE|O,cColor:3:55.86.35|O,W12|"
    XLS = XLS & "P,$IntE|P,cColor:3:55.86.35|P,W12|"
    XLS = XLS & "Q,$IntE|Q,cColor:3:55.86.35|Q,W12|"
    XLS = XLS & "R,BC|R,cBorder:Med|R,cColor:3:84.130.53|R,W06|"
    XLS = XLS & "S,SintE|S,cBorder:Med:End=T|S,cColor:3:84.130.53|S,W11|"
    XLS = XLS & "T,$IntE|T,cColor:3:84.130.53|T,W11|"
    XLS = XLS & "U,Sint|U,cColor:3:191.143.0|U,CondFormat:Range1.>599999:Color.RGB(255.255.0)|U,cBorder:Med|U,W10|"
    XLS = XLS & "V,Sint|V,cColor:3:191.143.0|V,W12|"
    XLS = XLS & "W,Sint|W,cColor:3:191.143.0|W,W12|"
    XLS = XLS & "X,$IntE|X,cColor:3:128.96.0|X,cBorder:Med|X,W10|"
    XLS = XLS & "Y,W06|"
    XLS = XLS & "Z,TL|Z,W11|"
    XLS = XLS & "AA,TC|AA,cColor:3:0.0.0|AA,W07|"
    XLS = XLS & "AB,TC|AB,W08|"
    XLS = XLS & "AC,TC|AC,|AC,cColor:3:0.0.0|AC,W15|"
    XLS = XLS & "AD,TL|AD,W25|"
    TitleX = "Bank Sales Report - Between: " & Date - 365 & " And: " & Date
    If sBankName <> "" Then TitleX = TitleX & " - " & sBankName
    MessageUser "Open", "Creating Bank Sales Report, That's Tab 1", "...Excel Will Open in a bit..."
    SendReport2Excel SQLz, "Sold Units", TitleX, True, XLS, , False, 1


Inventory:
    SQLz = "Select SourceCompany AS [Source Client]"                    '- Col:A
    SQLz = SQLz & " ,BankPartner AS Bank"                               '- Col:B
    SQLz = SQLz & " ,UnitVin AS VIN"                                    '- Col:C
    SQLz = SQLz & " ,UnitStatus AS [Unit Status]"                       '- Col:D
    SQLz = SQLz & " ,UnitDescShortVer AS [Unit Desc]"                   '- Col:E
    SQLz = SQLz & " ,PState"                                            '- Col:F
    SQLz = SQLz & " ,LeaseExp AS [Lease Expire]"                        '- Col:G
    SQLz = SQLz & " ,OffLeaseDate AS [Actual OffLease]"                 '- Col:H
    SQLz = SQLz & " ,Rstatus  AS [Inventory Status]"                    '- Col:I
    SQLz = SQLz & " ,DaysInInventory AS [Days IN]"                      '- Col:J
    SQLz = SQLz & " ,Residual AS RV"                                    '- Col:K
    SQLz = SQLz & " ,CurrBook As [Curr Book]"                           '- Col:L
    SQLz = SQLz & " ,TargetSalePrice AS [Target Price]"                 '- Col:M
    SQLz = SQLz & " ,[Expected GP] AS [Expected Gross Profit]"          '- Col:N
    SQLz = SQLz & " ,TotalExpense AS [RMKT Exp]"                        '- Col:O
    SQLz = SQLz & " ,0 as [Expected NET Profit]"                        '- Col:P
    SQLz = SQLz & " ,Odometer AS ODO"                                   '- Col:Q
    SQLz = SQLz & " ,Milesallowed AS [Miles Allowed]"                   '- Col:R
    SQLz = SQLz & " ,iif(OverMiles<0,0,OverMiles) AS [Over Miles]"      '- Col:S
    SQLz = SQLz & " ,OverMilesCost AS [Over Miles Cost]"                '- Col:T
    SQLz = SQLz & " ,Disposition AS Disp"                               '- Col:U
    SQLz = SQLz & " ,[Sch|Grp] AS [Lease Sch|Grp]"                      '- Col:V
    SQLz = SQLz & " ,UnitId AS AssetID"                                 '- Col:W
    SQLz = SQLz & " ,Unitnum"                                           '- Col:X
    SQLz = SQLz & " ,UnitSubType AS SubType"                            '- Col:Y
    SQLz = SQLz & " ,ParkedLocation As [Parked Location]"               '- Col:Z
    SQLz = SQLz & " ,BankShareYN AS [Bank Share Y/N]"                   '- Col:AA
    SQLz = SQLz & " ,EOption As [Exit Options]"                         '- Col:AB
    SQLz = SQLz & " ,(Odometer + ProjMiles) As [Proj ODO EOL]"          '- Col:AC
    SQLz = SQLz & " ,BuyOut"                                            '- Col:AD
    SQLz = SQLz & " FROM vw_RemarketingReportAllUnits "
    SQLz = SQLz & " WHERE UnitStatus IN ('A', 'TR', 'R','PS','P','ES','EH','EP','EI','EF','EW','EG','EM','EZ')"
    If iBankId > 0 Then SQLz = SQLz & " AND BankID = " & iBankId & " "
    SQLz = SQLz & " ORDER BY SourceCompany, BankPartner;"

    XLS = "|A,mTable:Name=Inventory|A,DD3|A,W12|"
    XLS = XLS & "B,TL|B,W10|"
    XLS = XLS & "C,TC|C,TX|C,W19|"
    XLS = XLS & "D,BC|D,cColor:3:0.0.0|D,cBorder:Med|D,W06|"
    XLS = XLS & "E,BC|E,W20|"
    XLS = XLS & "F,TC|F,W05.5|"
    XLS = XLS & "G,TC|G,W10|"
    XLS = XLS & "H,TC|H,W10|"
    XLS = XLS & "I,TC|I,W11|"
    XLS = XLS & "J,BC|J,cColor:3:0.0.0|J,W05|"
    XLS = XLS & "K,$IntE|K,W12|"
    XLS = XLS & "L,$IntE|L,cColor:3:0.0.0|L,cBorder:Med|L,W12|"
    XLS = XLS & "M,$IntE|M,cColor:3:55.86.35|M,W12|"
    XLS = XLS & "N,$IntE|N,cLabel:End=N:3:Expected Gross Profit|N,cColor:3:55.86.35|N,W12|"
    XLS = XLS & "O,$IntE|O,cColor:3:55.86.35|O,W10|"
    XLS = XLS & "P,$IntE|P,cColor:3:55.86.35|P,cFormula:FxAll=IF([@[Target Price]]>0,[@[Target Price]]-([@[RMKT Exp]] + [@[Curr Book]]),0)|P,W12|"
    XLS = XLS & "Q,Sint|Q,cColor:3:191.143.0|Q,CondFormat:Range1.>599999:Color.RGB(255.255.0)|Q,cBorder:Med|Q,W09|"
    XLS = XLS & "R,Sint|R,cColor:3:191.143.0|R,W09|"
    XLS = XLS & "S,Sint|S,cColor:3:191.143.0|S,W09|"
    XLS = XLS & "T,$IntE|T,cColor:3:128.96.0|T,cBorder:Med|T,W10|"
    XLS = XLS & "U,TC|U,W07|"
    XLS = XLS & "V,TL|W,W11|"
    XLS = XLS & "W,TX|W,cColor:3:0.0.0|W,cBorder:Med|W,W10|"
    XLS = XLS & "X,TC|X,W08|"
    XLS = XLS & "Y,TC|Y,cColor:3:0.0.0|Y,W15|"
    XLS = XLS & "Z,TL|Z,W25|"
    XLS = XLS & "AA,BC|AA,cBorder:Med|AB,W07|"
    XLS = XLS & "AB,TC|AB,W11|"
    XLS = XLS & "AC,Sint|AC,cColor:3:crimson|AC,CondFormat:Range1.>599999:Color.RGB(255.255.0)|AC,cBorder:Med|AC,W09|"
    XLS = XLS & "AD,$IntE|AD,W12|"
    TitleX = "Bank Inventory Report - " & Date
    If sBankName <> "" Then TitleX = TitleX & " - " & sBankName
    MessageUser "Open", "Creating Bank Inventory Report, Tab 2", "...Excel Will Open in a bit..."
    SendReport2Excel SQLz, "Inventory", TitleX, True, XLS, , False, 2, True
    MessageUser "Close"
End Function


Function RVRealizationReport(TitleX As String, DateFilter As String, AndWhereX As String)

    Dim TitlePreX As String
    XLS = ""
    SQLx = "SELECT "
    SQLx = SQLx & DateFilter & "Year"
    SQLx = SQLx & ", SourceCompany AS Client"
    SQLx = SQLx & ", Schedule"
    SQLx = SQLx & ", BankPartner as Assignee"
    SQLx = SQLx & ", UnitID"
    SQLx = SQLx & ", iif(isdate(ExtendedEOL),RU.[Disposition] & '-Extd', RU.[Disposition]) AS Disposition"
    SQLx = SQLx & ", UnitYr As MYear"
    SQLx = SQLx & ", UnitMake As Make"
    SQLx = SQLx & ", UnitModel As Model"
    SQLx = SQLx & ", UnitType As Type"
    SQLx = SQLx & ", BLCD"
    SQLx = SQLx & ", " & SplitString(SplitString(DateFilter, ",", 1), ")", 0) & ""
    SQLx = SQLx & ", LeaseTermPrimary as PTerm"
    SQLx = SQLx & ", DateDiff('m',BLCD,OffLeaseDate) as ExpiryAge"
    SQLx = SQLx & ", OEC"
    SQLx = SQLx & ", Residual"
    SQLx = SQLx & ", Residual/OEC As [RV%]"
    SQLx = SQLx & ", TotalExpense AS [RMKT Expense]"
    SQLx = SQLx & ", UnitSalePrice AS Sold"
    SQLx = SQLx & ", (UnitSalePrice-TotalExpense)/OEC As [Sold%]"
    SQLx = SQLx & ", BuyOut"
    SQLx = SQLx & ", BuyOut/OEC AS [BuyOut%]"
    SQLx = SQLx & ", (UnitSalePrice-TotalExpense)-Residual As [RV Gain/Loss]"
    SQLx = SQLx & ", (UnitSalePrice-TotalExpense)/Residual As [RV Gain/Loss%]"
    SQLx = SQLx & ", DaysInInventory AS InvDays"
    If DateFilter Like "*BuyOutDate*" Then SQLx = SQLx & "" Else SQLx = SQLx & ", BuyOutDate As BuyOut"
    SQLx = SQLx & ", InvoiceSent"
    If DateFilter Like "*InvoicePaid*" Then SQLx = SQLx & "" Else SQLx = SQLx & ", InvoicePaid"
    SQLx = SQLx & ", Odometer"
    SQLx = SQLx & ", Lstate as State"
    SQLx = SQLx & ", ExtendedEOL"
    SQLx = SQLx & ", UnitTermPenalty"
    SQLx = SQLx & ", CurrBook"
    SQLx = SQLx & " FROM RemarketingReportAllUnits As RU"
    SQLx = SQLx & " WHERE IsPLB=False "
    SQLx = SQLx & " AND MLOrig='FA'"
    SQLx = SQLx & AndWhereX
    TitlePreX = "RV Realization Unit Detail "
    XLS = "|A,DD|A,TC|E,TX|G,TC|"
    XLS = XLS & "M,cLabel:End=M:3:Pri" & vbNewLine & "Term|M,TC|"
    XLS = XLS & "N,cLabel:End=N:3:Age @" & vbNewLine & "Expiry|N,BC|"
    XLS = XLS & "O,$IntE|"
    XLS = XLS & "P,$IntE|"
    XLS = XLS & "Q,P2|"
    XLS = XLS & "R,cLabel:End=R:3:RMKT" & vbNewLine & "Exp|R,$IntE|"
    XLS = XLS & "S,cLabel:End=S:3:Sold" & vbNewLine & "Amt|S,$IntE|S,cColor:3:51.102.0|"       ' Green
    XLS = XLS & "T,cLabel:End=T:3:Sold vs" & vbNewLine & "OEC%|T,P2|T,cColor:3:51.102.0|"      ' Green
    XLS = XLS & "U,cLabel:End=U:3:BuyOut" & vbNewLine & "Amt|U,$IntE|U,cColor:3:222.132.0|"    ' MedOrange
    XLS = XLS & "V,cLabel:End=V:3:BO vs" & vbNewLine & "OEC%|V,P2|V,cColor:3:222.132.0|"       ' MedOrange
    XLS = XLS & "W,$E|X,P2|Y,BC|AB,SR|AC,BC|AE,$IntE|AF,$IntE|"
    XLS = XLS & "AE,cLabel:End=AE:3:Unit Term" & vbNewLine & "Penalty|AE,cColor:3:192.0.0|"    ' MedRed
    MessageUser "Open", "Unit Level, Tab 3", "...This will take just a moment..."
    SendReport2Excel SQLx, "Unit Detail", TitlePreX & TitleX, True, XLS, , False, 1

    XLS = ""
    SQLx = "SELECT "
    SQLx = SQLx & DateFilter & "Year"
    SQLx = SQLx & ", DatePart('yyyy',RU.BLCD) AS BLCD"
    SQLx = SQLx & ", Count(UnitID) As Units"
    SQLx = SQLx & ", iif(isdate(ExtendedEOL),RU.[Disposition] & '-Extd', RU.[Disposition]) AS Disposition"
    SQLx = SQLx & ", Sum(OEC) as [OEC Total]"
    SQLx = SQLx & ", Avg(OEC) as [OEC Avg]"
    SQLx = SQLx & ", Sum(Residual) as [RV Total]"
    SQLx = SQLx & ", Avg(Residual) as [RV Avg]"
    SQLx = SQLx & ", Sum(Residual)/Sum(OEC) As [RV%]"
    SQLx = SQLx & ", Sum(TotalExpense) AS [RMKT Expense]"
    SQLx = SQLx & ", Sum(UnitSalePrice) AS Sold"
    SQLx = SQLx & ", Avg(UnitSalePrice) AS [Sold Avg]"
    SQLx = SQLx & ", (Sum(UnitSalePrice)-Sum(TotalExpense))/Sum(OEC) As [RMKT%]"
    SQLx = SQLx & ", (Sum(UnitSalePrice)-Sum(TotalExpense))-Sum(Residual) As [RV Gain/Loss]"
    SQLx = SQLx & ", (Sum(UnitSalePrice)-Sum(TotalExpense))/Sum(Residual) As [RV Gain/Loss%]"
    SQLx = SQLx & ", Round(Avg([DaysInInventory]),0) AS InvDays"
    SQLx = SQLx & " FROM RemarketingReportAllUnits As RU"
    SQLx = SQLx & " WHERE IsPLB=False "
    SQLx = SQLx & " AND MLOrig='FA'"
    SQLx = SQLx & AndWhereX
    SQLx = SQLx & " Group By " & SplitString(DateFilter, " AS", 0) & ", iif(isdate(ExtendedEOL),RU.[Disposition] & '-Extd', RU.[Disposition]),DatePart('yyyy',RU.BLCD)"
    TitlePreX = "RV Realization Summary "
    XLS = "|A,DD|A,TC|C,SE|C,TC|E,$IntE|F,$IntE|G,$IntE|H,$IntE|I,P2|J,$IntE|K,$IntE|L,$IntE|M,P2|N,$IntE|O,P2|P,SInt|"
    MessageUser "Update", "Realization Summary, Tab 2", "...This will take just a moment..."
    SendReport2Excel SQLx, "Realization", TitlePreX & TitleX, True, XLS, , False, 2

    XLS = ""
    SQLx = "SELECT "
    SQLx = SQLx & DateFilter & "Year"
    SQLx = SQLx & ", DatePart('yyyy',RU.BLCD) AS BLCD"
    SQLx = SQLx & ", Count(UnitID) AS Units"
    SQLx = SQLx & ", iif(isdate(ExtendedEOL),RU.[Disposition] & '-Extd', RU.[Disposition]) AS Disposition"
    SQLx = SQLx & ", Sum(UnitSalePrice) AS [Sold Total]"
    SQLx = SQLx & ", Sum(UnitSalePrice) AS [Sold Avg]"
    SQLx = SQLx & ", Sum(TotalExpense) AS [Expense Total]"
    SQLx = SQLx & ", Round(Avg([TotalExpense]),0) AS [Expense Avg]"
    SQLx = SQLx & ", Round(Avg([DaysInInventory]),0) AS [InvDays Avg]"
    SQLx = SQLx & " FROM RemarketingReportAllUnits As RU"
    SQLx = SQLx & " WHERE IsPLB=False "
    SQLx = SQLx & " AND MLOrig='FA'"
    SQLx = SQLx & AndWhereX
    SQLx = SQLx & " Group By " & SplitString(DateFilter, " AS", 0) & ", iif(isdate(ExtendedEOL),RU.[Disposition] & '-Extd', RU.[Disposition]),DatePart('yyyy',RU.BLCD)"
    XLS = "|A,DD|A,TC|C,SE|C,TC|E,$E|F,$IntE|G,$IntE|H,$IntE|I,SInt|"
    TitlePreX = "Results Summary "
    MessageUser "Update", "Creating Summary, That's Tab 1", "...Excel Will Open in a bit..."
    SendReport2Excel SQLx, "Result Summary", TitlePreX & TitleX, True, XLS, , True, 3, True

    MessageUser "Close"
End Function

Public Function UnSellAUnit(iAssetID As Long, Optional iBuyerOrderID As Long)
    Committ = MsgBox("You're About 'Un-Sell' AssetID: " & iAssetID & vbNewLine & "Do You Agree?" & vbNewLine & vbNewLine & "This cannot be reversed", vbQuestion + vbYesNo + vbCritical, "Un-Sold A Unit")
    If Committ = vbNo Then Exit Function
     SQLx = "Update RMKT Set "
     SQLx = SQLx & "InvoicePaidDate = Null"
     SQLx = SQLx & ",InvoiceRequestDate = Null"
     SQLx = SQLx & ",InvoiceSentDate = Null"
     SQLx = SQLx & ",SoldPrice = Null"
     SQLx = SQLx & ",BuyerRef = Null"
     SQLx = SQLx & ",Wholesale = Null"
     SQLx = SQLx & ",RmktBOSSent = Null"
     SQLx = SQLx & ",RmktDeposit = Null"
     SQLx = SQLx & ",RmktDepdt = Null"
     SQLx = SQLx & ",RmkSm = Null"
     SQLx = SQLx & ",GainLoss = Null"
     SQLx = SQLx & ",InvoiceRequestNum = Null"
     SQLx = SQLx & " Where UnitREF = " & iAssetID
     DoCmd.SetWarnings False
     DoCmd.RunSQL (SQLx)
     SQLx = "Update Units set UnitStatus = 'R' Where UnitID = " & iAssetID
     DoCmd.RunSQL (SQLx)

     SQLx = "Delete From RMKT_BuyersOrders Where UnitID = " & iAssetID & " AND BuyersOrderID=" & iBuyerOrderID & ";"
     DoCmd.SetWarnings False
     If iBuyerOrderID > 0 Then DoCmd.RunSQL (SQLx)
     DoCmd.SetWarnings True
     If iBuyerOrderID > 0 Then Exit Function

     Committ = MsgBox("Do you want to Remove this Unit: " & iAssetID & " from ALL Buyers Orders" & vbNewLine & "Yes or No?" & vbNewLine & vbNewLine & "This cannot be reversed", vbQuestion + vbYesNo + vbCritical, "Un-Sold A Unit")
     If Committ = vbNo Then Exit Function
     SQLx = "Delete From RMKT_BuyersOrders Where UnitID = " & iAssetID
     DoCmd.SetWarnings False
     DoCmd.RunSQL (SQLx)
     DoCmd.SetWarnings True

End Function
Public Function AreUnitsInRMKT() As Integer
    Set RST = CurrentDb.OpenRecordset("SELECT RMKT_MMod_Work.UnitID FROM RMKT_MMod_Work LEFT JOIN Rmkt ON RMKT_MMod_Work.UnitID = Rmkt.UnitRef WHERE (((Rmkt.UnitRef) Is Null))")
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    AreUnitsInRMKT = RST.RecordCount
    RST.close
    Set RST = Nothing
End Function
Public Function InsertMassRMKTUnits()
    SQLx = "INSERT INTO Rmkt ( UnitRef, Discount ) "
    SQLx = SQLx & "SELECT W.UnitID, 0 AS Discount "
    SQLx = SQLx & "FROM RMKT_MMod_Work AS W LEFT JOIN Rmkt ON W.UnitID = Rmkt.UnitRef "
    SQLx = SQLx & "WHERE (((Rmkt.UnitRef) Is Null));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLx)
    DoCmd.SetWarnings True
End Function
Public Function IsUnitInRMKT(iUnitID As Long) As Integer
    IsUnitInRMKT = DCount("UnitREf", "RMKT", "UnitRef=" & iUnitID)
End Function
Public Function InsertRMKTUnit(iUnitID)
    SQLx = "INSERT INTO Rmkt ( UnitRef, Discount ) "
    SQLx = SQLx & "SELECT " & iUnitID & ", 0 AS Discount "
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLx)
    DoCmd.SetWarnings True
End Function
