Operation =1
Option =0
Where ="(((Mileage.UnitID) Is Not Null))"
Having ="(((lookup_groupTypes.grpClass)<>\"trailer\") AND ((Clients.clientobc)<>\"none\")"
    " AND ((Units.UnitExckPortRpt)=False))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Mileage"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="Mileage.UnitID"
    Alias ="LastOfMHDate"
    Expression ="Max(Mileage.LogActivityDate)"
    Alias ="Odometer"
    Expression ="Max(Mileage.OdometerEnd)"
    Expression ="lookup_groupTypes.grpClass"
    Expression ="Clients.clientobc"
    Expression ="Units.UnitExckPortRpt"
    Alias ="UnitRef"
    Expression ="Units.UnitID"
End
Begin Joins
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Units"
    RightTable ="Mileage"
    Expression ="Units.UnitID = Mileage.UnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Mileage.UnitID"
    Flag =0
End
Begin Groups
    Expression ="Mileage.UnitID"
    GroupLevel =0
    Expression ="lookup_groupTypes.grpClass"
    GroupLevel =0
    Expression ="Clients.clientobc"
    GroupLevel =0
    Expression ="Units.UnitExckPortRpt"
    GroupLevel =0
    Expression ="Units.UnitID"
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
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Clients.clientobc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.UnitExckPortRpt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Odometer"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfMHDate"
        dbInteger "ColumnWidth" ="4020"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.OdometerEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OdometerEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfOdometerEnd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =369
    Right =1279
    Bottom =1015
    Left =-1
    Top =-1
    Right =1226
    Bottom =341
    Left =15
    Top =0
    ColumnsShown =543
    Begin
        Left =5
        Top =2
        Right =149
        Bottom =146
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =167
        Top =6
        Right =311
        Bottom =150
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =310
        Top =0
        Right =454
        Bottom =144
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =475
        Top =-12
        Right =619
        Bottom =132
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =654
        Top =-2
        Right =798
        Bottom =142
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =845
        Top =39
        Right =989
        Bottom =321
        Top =0
        Name ="Mileage"
        Name =""
    End
    Begin
        Left =662
        Top =166
        Right =806
        Bottom =310
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
