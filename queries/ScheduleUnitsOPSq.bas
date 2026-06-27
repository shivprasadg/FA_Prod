Operation =1
Option =0
Where ="(((Schedule.SchID)=[Forms]![ScheduleControl]![schid]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="docsreeferq"
    Name ="Last Title Entry"
End
Begin OutputColumns
    Alias ="Unit No"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitUnitNum2"
    Alias ="Asset #"
    Expression ="Units.UnitID"
    Expression ="Units.Unitvin"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & \" \" & [UnitType] & IIf([Axle]=\"tbd\",Null"
        ",\" \" & [axle])"
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
    Alias ="RefrigU Make"
    Expression ="docsreeferq.ChildMake"
    Alias ="RefrigU Model"
    Expression ="docsreeferq.ChildClass"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.unitstatus"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LocationStreet"
    Alias ="MSO Copy Rec"
    Expression ="Units.unitmsocopy"
    Alias ="Title App Sent"
    Expression ="Units.UnitTASent"
    Alias ="Title"
    Expression ="[Last Title Entry].Action"
    Expression ="Units.UnitPlateNum"
    Alias ="InserviceDate"
    Expression ="Units.UnitInServ"
    Alias ="Date"
    Expression ="[Last Title Entry].LastOfUnitTitleLogDate"
    Alias ="UnitCost"
    Expression ="Units.UnitFinalCost"
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
    LeftTable ="Units"
    RightTable ="Last Title Entry"
    Expression ="Units.UnitID = [Last Title Entry].UnitRef"
    Flag =2
End
Begin OrderBy
    Expression ="Units.UnitID"
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
        dbInteger "ColumnWidth" ="3105"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="3825"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit No"
        dbInteger "ColumnWidth" ="2490"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1890"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="1890"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Co_Lessee"
        dbInteger "ColumnWidth" ="3420"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease No"
        dbInteger "ColumnWidth" ="1470"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1560"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbInteger "ColumnOrder" ="15"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbInteger "ColumnWidth" ="1485"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="docsreeferq.[Serial Number]"
        dbInteger "ColumnOrder" ="11"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbInteger "ColumnWidth" ="1350"
        dbInteger "ColumnOrder" ="17"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbInteger "ColumnWidth" ="1410"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="3390"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbInteger "ColumnWidth" ="1755"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title App Sent"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MSO Copy Rec"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RefrigU Make"
        dbInteger "ColumnWidth" ="1665"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigU Model"
        dbInteger "ColumnWidth" ="1875"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Asset #"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InserviceDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1177
    Bottom =793
    Left =-1
    Top =-1
    Right =1145
    Bottom =161
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =14
        Top =26
        Right =158
        Bottom =170
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =193
        Top =15
        Right =337
        Bottom =159
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =364
        Top =16
        Right =508
        Bottom =160
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =532
        Top =17
        Right =676
        Bottom =161
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =702
        Top =15
        Right =846
        Bottom =159
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =902
        Top =152
        Right =1046
        Bottom =292
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =900
        Top =-2
        Right =1044
        Bottom =142
        Top =0
        Name ="docsreeferq"
        Name =""
    End
    Begin
        Left =206
        Top =168
        Right =461
        Bottom =312
        Top =0
        Name ="Last Title Entry"
        Name =""
    End
End
