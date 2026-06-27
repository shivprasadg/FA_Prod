Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposalUnitCountDistinct"
End
Begin OutputColumns
    Expression ="BuyersProposalUnitCountDistinct.SalesPersonID"
    Alias ="UnitCount"
    Expression ="Count(BuyersProposalUnitCountDistinct.UnitID)"
End
Begin Groups
    Expression ="BuyersProposalUnitCountDistinct.SalesPersonID"
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
        dbText "Name" ="BuyersProposalUnitCountDistinct.SalesPersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unitcount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1482
    Bottom =1038
    Left =-1
    Top =-1
    Right =1450
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =153
        Top =205
        Right =297
        Bottom =349
        Top =0
        Name ="BuyersProposalUnitCountDistinct"
        Name =""
    End
End
