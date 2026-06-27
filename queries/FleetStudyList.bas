Operation =1
Option =0
Begin InputTables
    Name ="FleetStudyImport"
End
Begin OutputColumns
    Expression ="FleetStudyImport.VIN"
    Expression ="FleetStudyImport.Unitnumber"
    Expression ="FleetStudyImport.UnitYear"
    Expression ="FleetStudyImport.UnitMake"
    Expression ="FleetStudyImport.UnitModel"
    Expression ="FleetStudyImport.Type"
    Expression ="FleetStudyImport.Location"
End
Begin OrderBy
    Expression ="FleetStudyImport.VIN"
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
        dbText "Name" ="FleetStudyImport.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.UnitYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.Unitnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImport.Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =129
    Top =193
    Right =1343
    Bottom =974
    Left =-1
    Top =-1
    Right =1182
    Bottom =447
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =407
        Bottom =482
        Top =0
        Name ="FleetStudyImport"
        Name =""
    End
End
