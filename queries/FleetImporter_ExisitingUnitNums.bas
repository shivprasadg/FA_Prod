Operation =1
Option =0
Where ="(((Units.ClientID)=[Forms]![FleetStudy_Importer]![ClientID]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.ClientGroupId"
    Expression ="Units.ClientID"
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
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =163
    Top =14
    Right =1304
    Bottom =995
    Left =-1
    Top =-1
    Right =1109
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =120
        Top =66
        Right =537
        Bottom =527
        Top =0
        Name ="Units"
        Name =""
    End
End
