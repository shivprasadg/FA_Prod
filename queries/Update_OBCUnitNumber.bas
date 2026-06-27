Operation =1
Option =0
Where ="(((Units.ObcUnitNumber) Is Null))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.ClientGroupId"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.ObcUnitNumber"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =337
    Top =46
    Right =1809
    Bottom =1044
    Left =-1
    Top =-1
    Right =1440
    Bottom =698
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =133
        Top =141
        Right =384
        Bottom =645
        Top =0
        Name ="Units"
        Name =""
    End
End
