Operation =1
Option =0
Where ="(((SchGrp.LeaseType)<>\"loan\") AND ((Units.unitstatus)=\"A\" Or (Units.unitstat"
    "us)=\"B\" Or (Units.unitstatus)=\"C\") AND ((Units.ClientGroupId)<>435))"
Having ="(((Clients.clientID)=Forms!MgmtMenu!SelCP) And ((Units.UnitClassType)<>\"Yard Tr"
    "actor\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Alias ="Class"
    Expression ="Units.UnitClassType"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
    Expression ="Clients.clientID"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitClassType"
    Flag =0
End
Begin Groups
    Expression ="Units.UnitClassType"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
End
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
        dbText "Name" ="CountOfUnitID"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[Client Group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =275
    Top =23
    Right =1533
    Bottom =669
    Left =-1
    Top =-1
    Right =1226
    Bottom =79
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =895
        Top =12
        Right =1039
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =31
        Top =-11
        Right =175
        Bottom =133
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =212
        Top =-10
        Right =356
        Bottom =134
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =381
        Top =0
        Right =525
        Bottom =144
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =544
        Top =3
        Right =682
        Bottom =139
        Top =0
        Name ="Units"
        Name =""
    End
End
