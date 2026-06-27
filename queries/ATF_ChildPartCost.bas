Operation =1
Option =0
Begin InputTables
    Name ="ATFChildPartSpec"
End
Begin OutputColumns
    Expression ="ATFChildPartSpec.ATFUnitSpecID"
    Alias ="ChildPartCost"
    Expression ="Sum(ATFChildPartSpec.ChildPartCost)"
    Alias ="Parts"
    Expression ="Count(ATFChildPartSpec.ATFChildSpecID)"
    Expression ="ATFChildPartSpec.ATFID"
End
Begin Groups
    Expression ="ATFChildPartSpec.ATFUnitSpecID"
    GroupLevel =0
    Expression ="ATFChildPartSpec.ATFID"
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
        dbText "Name" ="SumOfChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ATFChildSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfATFChildSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parts"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =516
    Top =96
    Right =1764
    Bottom =919
    Left =-1
    Top =-1
    Right =1230
    Bottom =546
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =210
        Top =101
        Right =354
        Bottom =245
        Top =0
        Name ="ATFChildPartSpec"
        Name =""
    End
End
