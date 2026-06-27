Operation =1
Option =0
Where ="(((Units.UnitID)=[forms]![UnitsForm]![unitid]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Mileage"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitID"
    Expression ="SchGrp.LeaseType"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.unitvin"
    Alias ="Expr1"
    Expression ="Mileage.MHDate"
    Alias ="Expr2"
    Expression ="Mileage.PerFuelG"
    Alias ="Expr3"
    Expression ="Mileage.PerIdleG"
    Alias ="Expr4"
    Expression ="Mileage.PerMi"
    Alias ="Odometer"
    Expression ="Mileage.AquiMI"
    Alias ="Expr5"
    Expression ="Mileage.MovingMPG"
    Alias ="Expr6"
    Expression ="Mileage.OverAllMPG"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Mileage"
    Expression ="Units.UnitID = Mileage.UnitID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
End
Begin OrderBy
    Expression ="Mileage.MHDate"
    Flag =0
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.MHDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerFuelG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerIdleG"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Mileage.MovingMPG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.OverAllMPG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerMi"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =272
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
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1200
        Top =12
        Right =1344
        Bottom =156
        Top =0
        Name ="Mileage"
        Name =""
    End
End
