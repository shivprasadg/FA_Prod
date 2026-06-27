Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Bulk_Update_WorkTable"
End
Begin OutputColumns
    Alias ="Asset"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Alias ="VendorId"
    Expression ="Units.VendorREF"
    Alias ="VIN"
    Expression ="Units.Unitvin"
    Expression ="Units.UnitBuildDate"
    Expression ="Units.UnitEstDelDate"
    Expression ="Units.UnitActualDelDate"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitInServ"
    Expression ="Units.UnitDeliveryLocationRef"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitType"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Bulk_Update_WorkTable"
    Expression ="Units.UnitID = Bulk_Update_WorkTable.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitID"
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
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Asset"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[unittype]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =93
    Top =47
    Right =1721
    Bottom =862
    Left =-1
    Top =-1
    Right =1596
    Bottom =332
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =218
        Top =65
        Right =410
        Bottom =412
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =64
        Top =24
        Right =182
        Bottom =167
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
End
