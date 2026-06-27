Operation =1
Option =0
Having ="(((MstrLease.MLOrig)=\"fa\") AND ((IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid])"
    ",[schfaid],[sgrpfaid]),[unitfaid]))=\"13-545\"))"
Begin InputTables
    Name ="Locations"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Alias ="CountOfClientShNm"
    Expression ="Count(Clients.ClientShNm)"
    Expression ="MstrLease.MLOrig"
    Alias ="FAID"
    Expression ="IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]),[unitfaid])"
    Expression ="Units.VendorREF"
    Alias ="Main Unit"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [axle]"
    Expression ="Locations.LocationName"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LAddrLine3"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Units.unitfinalcost"
    Expression ="Locations.LZip"
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
    LeftTable ="Locations"
    RightTable ="Units"
    Expression ="Locations.LocationID = Units.UnitDeliveryLocationRef"
    Flag =3
End
Begin Groups
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]),[unitfaid])"
    GroupLevel =0
    Expression ="Units.VendorREF"
    GroupLevel =0
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [axle]"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
    GroupLevel =0
    Expression ="Locations.LAddrLine1"
    GroupLevel =0
    Expression ="Locations.LAddrLine2"
    GroupLevel =0
    Expression ="Locations.LAddrLine3"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Units.unitfinalcost"
    GroupLevel =0
    Expression ="Locations.LZip"
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
        dbText "Name" ="MstrLease.MLOrig"
        dbInteger "ColumnWidth" ="1080"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="4185"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbInteger "ColumnWidth" ="2055"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt4"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1560"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbInteger "ColumnWidth" ="1200"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitfinalcost"
        dbInteger "ColumnWidth" ="1590"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfClientShNm"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Main Unit"
        dbInteger "ColumnWidth" ="3810"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =436
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =1164
        Top =179
        Right =1308
        Bottom =323
        Top =0
        Name ="Locations"
        Name =""
    End
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
        Left =794
        Top =12
        Right =996
        Bottom =324
        Top =0
        Name ="Units"
        Name =""
    End
End
