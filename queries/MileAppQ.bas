Operation =1
Option =0
Where ="(((Clients.ClientShNm) Like \"hanna*\"))"
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
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Mileage.*"
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
    LeftTable ="Units"
    RightTable ="Mileage"
    Expression ="Units.UnitID = Mileage.UnitID"
    Flag =2
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="Units.UnitUnitNum"
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
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.mileageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.AquiMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.MileageDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.AqiHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.HoursDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.RecDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.IdleTime1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.IdleType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.IdleTime2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.IdleType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerMi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerFuelG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerIdleG"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerSpMin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerSpDist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerAvgStop"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.PerStopMin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Per M&R$"
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
    Bottom =238
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-12
        Top =3
        Right =132
        Bottom =147
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =171
        Top =8
        Right =315
        Bottom =152
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =332
        Top =7
        Right =476
        Bottom =151
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =493
        Top =5
        Right =637
        Bottom =149
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =647
        Top =1
        Right =791
        Bottom =145
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =833
        Top =7
        Right =977
        Bottom =151
        Top =0
        Name ="Mileage"
        Name =""
    End
End
