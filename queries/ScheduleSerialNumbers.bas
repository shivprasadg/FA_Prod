Operation =1
Option =0
Begin InputTables
    Name ="Schedule_Group_WorkTable"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="ChildSerq"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.unitunitnum"
    Alias ="VIN4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitEngineSerial"
    Expression ="Units.UnitTransSerial"
    Expression ="Units.UnitFASerial"
    Expression ="Units.UnitRAFSerial"
    Expression ="Units.UnitRARSerial"
    Expression ="Locations.LState"
    Expression ="ChildSerq.ChildMake"
    Expression ="ChildSerq.ChildClass"
    Expression ="ChildSerq.ChildType"
    Expression ="ChildSerq.ChildVIN"
    Expression ="ChildSerq.REFENG"
    Expression ="ChildSerq.COMPRESSS"
    Expression ="ChildSerq.MICROPROC"
    Expression ="ChildSerq.ELECDISP"
    Expression ="ChildSerq.[ARB ID]"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
    Expression ="Schedule.SchID"
    Expression ="Units.unitunitnum2"
    Expression ="Units.axle"
    Expression ="Units.unitsubtype"
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
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildSerq"
    Expression ="Units.UnitID = ChildSerq.UnitREF"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule_Group_WorkTable"
    RightTable ="Schedule"
    Expression ="Schedule_Group_WorkTable.ScheduleId = Schedule.SchID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Right([unitvin],6)"
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitsubtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.MICROPROC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.ELECDISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.REFENG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.[ARB ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildSerq.COMPRESSS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =404
    Top =157
    Right =1984
    Bottom =782
    Left =-1
    Top =-1
    Right =1548
    Bottom =257
    Left =109
    Top =0
    ColumnsShown =539
    Begin
        Left =-61
        Top =12
        Right =83
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =131
        Top =12
        Right =275
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =323
        Top =12
        Right =467
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =515
        Top =12
        Right =659
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =707
        Top =12
        Right =851
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =899
        Top =12
        Right =1043
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =1091
        Top =12
        Right =1235
        Bottom =156
        Top =0
        Name ="ChildSerq"
        Name =""
    End
    Begin
        Left =325
        Top =169
        Right =469
        Bottom =313
        Top =0
        Name ="Schedule_Group_WorkTable"
        Name =""
    End
End
