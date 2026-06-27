Operation =1
Option =0
Where ="(((Schedule.SchID)=[forms]![schedules_sub]![schid]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="docsreeferq"
End
Begin OutputColumns
    Alias ="Unit No"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitUnitNum2"
    Expression ="Units.Unitvin"
    Alias ="Desc"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType] & \" \" & IIf([Axle]=\"tbd\",Null,[axle])"
    Expression ="Schedule.SchNo"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="Co_Lessee"
    Expression ="Clients.clientCoNm2"
    Alias ="Lease No"
    Expression ="MstrLease.MLNo"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Expression ="docsreeferq.[Serial Number]"
    Alias ="RUnitMake"
    Expression ="docsreeferq.ChildMake"
    Alias ="RUnitModel"
    Expression ="docsreeferq.ChildClass"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.unitstatus"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LocationStreet"
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
    Flag =2
    LeftTable ="Units"
    RightTable ="docsreeferq"
    Expression ="Units.UnitID = docsreeferq.UnitREF"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitUnitNum"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
    Expression ="Units.UnitUnitNum"
    Flag =0
    Expression ="Right([unitvin],6)"
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
        dbText "Name" ="Units.Unitvin"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Unit No"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1485"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="3300"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Co_Lessee"
        dbInteger "ColumnWidth" ="3120"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Lease No"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="docsreeferq.[Serial Number]"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="RUnitMake"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="RUnitModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="3390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =225
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =95
        Right =181
        Bottom =239
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =277
        Top =90
        Right =421
        Bottom =234
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =436
        Top =3
        Right =580
        Bottom =147
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =613
        Top =-4
        Right =757
        Bottom =140
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =800
        Top =-12
        Right =944
        Bottom =132
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1042
        Top =110
        Right =1186
        Bottom =254
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =951
        Top =6
        Right =1095
        Bottom =150
        Top =0
        Name ="docsreeferq"
        Name =""
    End
End
