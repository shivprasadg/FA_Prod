Operation =1
Option =0
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.unitmake"
    Expression ="lookup_groupTypes.grpClass"
    Alias ="CountOfUnitID"
    Expression ="Count(Units.UnitID)"
    Expression ="Units.UnitSubType"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.unitmake"
    GroupLevel =0
    Expression ="lookup_groupTypes.grpClass"
    GroupLevel =0
    Expression ="Units.UnitSubType"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitSubType]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-196
    Top =58
    Right =1062
    Bottom =704
    Left =-1
    Top =-1
    Right =1226
    Bottom =347
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =164
        Top =7
        Right =308
        Bottom =148
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =356
        Top =12
        Right =500
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =548
        Top =12
        Right =777
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =799
        Top =9
        Right =943
        Bottom =153
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
