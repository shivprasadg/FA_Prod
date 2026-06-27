dbMemo "SQL" ="SELECT SourceCompany, BankPartner AS Bank, [Sch|Grp] AS [Lease Sch|Grp], UnitId "
    "AS AssetID, Unitnum, UnitVin AS VIN, UnitStatus AS [Unit Status], UnitYr AS Myea"
    "r, UnitMake AS Make, UnitModel AS Model, UnitType AS Type, UnitSubType AS SubTyp"
    "e, Axle, ParkedLocation AS [Parked Location], PState, ProjEOL AS [Projected Offl"
    "ease], OffLeaseDate AS [Actual OffLease], RStatus AS InvStatus, DaysInInventory "
    "AS [Days IN], Residual AS RV, CurrBook, TargetSalePrice AS [Target Price], [Expe"
    "cted GP], TotalExpense AS [RMKT Exp], Odometer AS ODO, Milesallowed AS [Miles Al"
    "lowed], OverMiles AS [Over|Under Miles], OverMilesCost AS [Over Miles Cost], Dis"
    "position AS Disp, BuyOut, ClientID, BankID\015\012FROM vw_RemarketingReportAllUn"
    "its\015\012WHERE (((UnitStatus) IN ('A', 'TR', 'R', 'P')) AND ((BankID) <> 4321)"
    ") OR (((UnitStatus) LIKE '*E*') AND ((BankID) <> 4321))\015\012ORDER BY SourceCo"
    "mpany, BankPartner;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Over|Under Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RemarketingReportAllUnits.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Sch|Grp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Myear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OffLease"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Over Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Over Miles Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days IN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrBook"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Target Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Expected GP]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT Exp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Projected Offlease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Miles Allowed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Disp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Actual OffLease"
        dbLong "AggregateType" ="-1"
    End
End
