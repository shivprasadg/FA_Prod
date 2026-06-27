dbMemo "SQL" ="SELECT Units.unitstatus, Clients.clientID, MstrLease.mlorig, Clients.clientshnm,"
    " MstrLease.mlno, Schedule.schno, Units.unitvin, Units.unitunitnum, Units.unityr,"
    " Units.unitmake, Units.unitmodel, Units.unittype, Rmkt.parkdt, Rmkt.BuyOutDate, "
    "lookup_GroupMakes.MakeAbrev, SchGrp.Assignee_Bank, Rmkt.BuyOutCost, Rmkt.SoldPri"
    "ce, Rmkt.RmktTitleSent, Rmkt.RmktBOSSent, Clients_1.ClientShNm, ChildPartSummary"
    "Query.SumOfchildFinalCost, DatePart(\"yyyy\",[InvoicePaidDate]) AS YR, Rmkt.MiPa"
    "rked, Rmkt.HrParked, Rmkt.BuyerRef, Rmkt.RmktSPRAdj, Rmkt.RmktAdjNote, Rmkt.RmkS"
    "m, Clients_1.[WH?], IIf([adjbookvalue]>0,[adjbookvalue],[BuyOutCost]) AS Cost, R"
    "mkt.TargetSale, Rmkt.InvoicePaidDate, lookup_groupTypes.GrpTypAbrev, Rmkt.RmktDe"
    "posit, Rmkt.RmktDepdt, IIf(IsNull([InvoicePaidDate]),\"NotSold\",\"Sold\") AS st"
    "atussort, RmktVinDecodeT.[Engine Series]\015\012FROM ((((Clients INNER JOIN Mstr"
    "Lease ON Clients.clientID = MstrLease.Client) INNER JOIN Schedule ON MstrLease.M"
    "LID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID)"
    " INNER JOIN (((((Units LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef) INNER JOIN"
    " lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake) LEFT JOIN Cl"
    "ients AS Clients_1 ON Rmkt.BuyerRef = Clients_1.clientID) LEFT JOIN ChildPartSum"
    "maryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF) INNER JOIN lookup_gro"
    "upTypes ON Units.UnitType = lookup_groupTypes.groupType) ON SchGrp.SGrpID = Unit"
    "s.GroupREF) LEFT JOIN RmktVinDecodeT ON Rmkt.UnitRef = RmktVinDecodeT.UnitID\015"
    "\012WHERE (((Units.unitstatus)=\"S\" Or (Units.unitstatus)=\"K\" Or (Units.units"
    "tatus)=\"KS\" Or (Units.unitstatus)=\"R\" Or (Units.unitstatus)=\"K\") AND ((Cli"
    "ents_1.ClientShNm) Is Not Null) AND ((Rmkt.InvoicePaidDate) Is Not Null) AND ((C"
    "lients.ClientGroupID)<> 18))\015\012ORDER BY Clients.clientshnm, MstrLease.mlno,"
    " Schedule.schid;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="MstrLease.mlorig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Clients.clientshnm"
        dbInteger "ColumnWidth" ="3045"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="MstrLease.mlno"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Schedule.schno"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Rmkt.parkdt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="lookup_GroupMakes.MakeAbrev"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpAssignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSalePr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Rmkt.RmktTitleSent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Rmkt.RmktBOSSent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Clients_1.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="ChildPartSummaryQuery.SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="YR"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Rmkt.MiParked"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Rmkt.HrParked"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSPRAdj"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Rmkt.RmktAdjNote"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Clients_1.[WH?]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.GrpTypAbrev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="statussort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktVinDecodeT.[Engine Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
End
