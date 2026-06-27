dbMemo "SQL" ="SELECT RemarketingReportAllUnits.SourceCompany AS [Account Name], RemarketingRep"
    "ortAllUnits.Schedule AS [Lease Sch], RemarketingReportAllUnits.UnitGroup AS [Gro"
    "up], RemarketingReportAllUnits.UnitNum AS [Customer Unit#], RemarketingReportAll"
    "Units.UnitVIN AS VIN, RemarketingReportAllUnits.UnitStatus AS [FA Unit Status], "
    "RemarketingReportAllUnits.UnitYr AS [Year], RemarketingReportAllUnits.UnitMake A"
    "S Make, RemarketingReportAllUnits.UnitModel AS Model, RemarketingReportAllUnits."
    "UnitType AS TYPE, [ParkedAddress] & \" \" & [PCity] & \", \" & [PState] AS Locat"
    "ion, RemarketingReportAllUnits.ParkDateProjected AS [Projected Park Date], Remar"
    "ketingReportAllUnits.InspectedDate AS [Inspection Date], RemarketingReportAllUni"
    "ts.InspectionLinkLocal AS [Local Link Insp], RemarketingReportAllUnits.Inspectio"
    "nLinkWeb AS [Web Link Insp], RemarketingReportAllUnits.Odometer, RemarketingRepo"
    "rtAllUnits.FMV AS [FA FMV], RemarketingReportAllUnits.TargetSalePrice AS [Target"
    " Sale Price], RemarketingReportAllUnits.UnitID, RemarketingReportAllUnits.Client"
    "ID, RemarketingReportAllUnits.FMV, RemarketingReportAllUnits.GainLoss, Remarketi"
    "ngReportAllUnits.TargetSalePrice, RemarketingReportAllUnits.LocationID, Remarket"
    "ingReportAllUnits.PCity, RemarketingReportAllUnits.PState, RemarketingReportAllU"
    "nits.BuyerID, RemarketingReportAllUnits.UnitYr, RemarketingReportAllUnits.UnitMa"
    "ke, RemarketingReportAllUnits.UnitModel, RemarketingReportAllUnits.UnitType, Rem"
    "arketingReportAllUnits.UnitStatus, RemarketingReportAllUnits.TitleRecv, Remarket"
    "ingReportAllUnits.InvoiceSent, RemarketingReportAllUnits.DaysInAcct AS [Days Sin"
    "ce BuyOut], RemarketingReportAllUnits.BuyOutDate, RemarketingReportAllUnits.[SUF"
    " Returned], RemarketingReportAllUnits.InvoicePaid\015\012FROM RemarketingReportA"
    "llUnits;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Type"
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
        dbText "Name" ="Target Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Projected Park Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Web Link Insp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Local Link Insp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA Unit Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inspection Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Sch"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TitleRecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days Since BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.PCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.PState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.GainLoss"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationState"
        dbLong "AggregateType" ="-1"
    End
End
