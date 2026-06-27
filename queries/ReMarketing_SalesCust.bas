Operation =1
Option =0
Where ="(((Units.unitstatus)=\"s\" Or (Units.unitstatus)=\"ks\" Or (Units.unitstatus)=\""
    "rs\") AND ((Clients_1.ClientShNm) Is Not Null) AND ((Clients.ClientGroupID)<> 18"
    "))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Rmkt"
    Name ="lookup_GroupMakes"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="ChildPartSummaryQuery"
    Name ="lookup_groupTypes"
    Name ="UnitVinDetails"
End
Begin OutputColumns
    Expression ="Units.unitstatus"
    Expression ="Clients.clientID"
    Expression ="MstrLease.mlorig"
    Expression ="Clients.clientshnm"
    Expression ="MstrLease.mlno"
    Expression ="Schedule.schno"
    Expression ="Units.unitvin"
    Expression ="Units.unitunitnum"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Rmkt.parkdt"
    Expression ="Rmkt.BuyOutDate"
    Expression ="lookup_GroupMakes.MakeAbrev"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="Rmkt.BuyOutCost"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.RmktTitleSent"
    Expression ="Rmkt.RmktBOSSent"
    Expression ="Clients_1.ClientShNm"
    Expression ="ChildPartSummaryQuery.SumOfchildFinalCost"
    Alias ="YR"
    Expression ="DatePart(\"yyyy\",[InvoicePaidDate])"
    Expression ="Rmkt.MiParked"
    Expression ="Rmkt.HrParked"
    Expression ="Rmkt.BuyerRef"
    Expression ="Rmkt.RmktSPRAdj"
    Expression ="Rmkt.RmktAdjNote"
    Expression ="Rmkt.RmkSm"
    Expression ="Clients_1.[WH?]"
    Alias ="Cost"
    Expression ="IIf([adjbookvalue]>0,[adjbookvalue],[BuyOutCost])"
    Expression ="Rmkt.TargetSale"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="lookup_groupTypes.GrpTypAbrev"
    Expression ="Rmkt.RmktDeposit"
    Expression ="Rmkt.RmktDepdt"
    Alias ="statussort"
    Expression ="IIf([unitstatus]=\"r\" Or [unitstatus]=\"k\",\"Pending Sale\",\"Sold\")"
    Alias ="Engine Series"
    Expression ="UnitVinDetails.series"
End
Begin Parameters
    Name ="[forms]![Rmkt Menu]![FromDate]"
    Flag =8
    Name ="[forms]![Rmkt Menu]![ToDate]"
    Flag =8
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_GroupMakes"
    Expression ="Units.UnitMake = lookup_GroupMakes.groupMake"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Clients_1"
    Expression ="Rmkt.BuyerRef = Clients_1.clientID"
    Flag =2
    LeftTable ="Units"
    RightTable ="ChildPartSummaryQuery"
    Expression ="Units.UnitID = ChildPartSummaryQuery.UnitREF"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="UnitVinDetails"
    Expression ="Rmkt.UnitRef = UnitVinDetails.unitid"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientshnm"
    Flag =0
    Expression ="MstrLease.mlno"
    Flag =0
    Expression ="Schedule.schid"
    Flag =0
End
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
    End
    Begin
        dbText "Name" ="MstrLease.mlorig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientshnm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3045"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MstrLease.mlno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.schno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.parkdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_GroupMakes.MakeAbrev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktTitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktBOSSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartSummaryQuery.SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.MiParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.HrParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSPRAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktAdjNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.[WH?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Engine Series"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =85
    Right =1520
    Bottom =806
    Left =-1
    Top =-1
    Right =1446
    Bottom =354
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1008
        Top =126
        Right =1152
        Bottom =270
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =1200
        Top =12
        Right =1344
        Bottom =156
        Top =0
        Name ="lookup_GroupMakes"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
    Begin
        Left =1210
        Top =203
        Right =1354
        Bottom =347
        Top =0
        Name ="UnitVinDetails"
        Name =""
    End
End
