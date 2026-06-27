Operation =1
Option =0
Having ="(((Units.UnitAxleProperty) Is Not Null Or (Units.UnitAxleProperty)<>\"\"))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitAxleProperty"
End
Begin OrderBy
    Expression ="Units.UnitAxleProperty"
    Flag =0
End
Begin Groups
    Expression ="Units.UnitAxleProperty"
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
        dbText "Name" ="Units.UnitAxleProperty"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1218
    Bottom =1021
    Left =-1
    Top =-1
    Right =1186
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =104
        Top =186
        Right =593
        Bottom =583
        Top =0
        Name ="Units"
        Name =""
    End
End
