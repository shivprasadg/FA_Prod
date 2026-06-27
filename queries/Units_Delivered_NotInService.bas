Operation =1
Option =0
Where ="(((Units.UnitStatus)=\"A\") AND ((Units.UnitActualDelDate) Is Not Null) AND ((Un"
    "its.UnitInServ) Is Null))"
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Alias ="AsetId"
    Expression ="Units.UnitID"
    Alias ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Alias ="COABackDate"
    Expression ="Units.UnitCOABack"
    Alias ="ModelYear"
    Expression ="Units.unitYr"
    Expression ="Units.UnitActualDelDate"
    Expression ="Units.UnitInServ"
    Expression ="Units.ReportStartDate"
    Expression ="Units.ReportEndDate"
    Alias ="DaysDelivered"
    Expression ="Date()-[UnitActualDelDate]"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Clients"
    Expression ="MstrLease.Client = Clients.clientID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
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
        dbText "Name" ="Units.ReportStartDate"
        dbInteger "ColumnWidth" ="2040"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COABackDate"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AsetId"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysDelivered"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2355"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =74
    Top =31
    Right =1950
    Bottom =885
    Left =-1
    Top =-1
    Right =1844
    Bottom =610
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =204
        Top =58
        Right =397
        Bottom =468
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =32
        Top =499
        Right =176
        Bottom =643
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =27
        Top =208
        Right =171
        Bottom =344
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =28
        Top =347
        Right =172
        Bottom =491
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =23
        Top =59
        Right =167
        Bottom =203
        Top =0
        Name ="Clients"
        Name =""
    End
End
