Operation =4
Option =0
Where ="(((Units.UnitModel) Like \"CASCADIA*\"))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.UnitModel"
    Expression ="Propercase([UnitModel],1)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query24].[diff]=\"Diff\")"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="diff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDetails.model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDetails.vin"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDetails.[make]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.clientgroupid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =126
    Top =11
    Right =1287
    Bottom =949
    Left =-1
    Top =-1
    Right =1129
    Bottom =621
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =155
        Top =85
        Right =360
        Bottom =438
        Top =0
        Name ="Units"
        Name =""
    End
End
