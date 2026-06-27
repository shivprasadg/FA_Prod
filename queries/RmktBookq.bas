Operation =1
Option =0
Where ="(((Units.unitstatus)=\"s\"))"
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
End
Begin OutputColumns
    Expression ="Units.unitstatus"
    Expression ="MstrLease.mlorig"
    Expression ="Clients.clientshnm"
    Expression ="MstrLease.mlno"
    Expression ="Schedule.schno"
    Expression ="Units.unitvin"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
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
    Expression ="Clients.[WH?]"
    Expression ="lookup_groupTypes.grpClass"
    Expression ="Clients_1.clientCompanyName"
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
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
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
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.mlorig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientshnm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.mlno"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Rmkt.RmktSalePr"
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
        dbText "Name" ="Clients.[WH?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitSubtype"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-453
    Top =44
    Right =805
    Bottom =690
    Left =-1
    Top =-1
    Right =1226
    Bottom =301
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
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =1009
        Top =10
        Right =1153
        Bottom =154
        Top =0
        Name ="lookup_GroupMakes"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =669
        Top =247
        Right =813
        Bottom =391
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
    Begin
        Left =1045
        Top =187
        Right =1189
        Bottom =331
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
