Operation =1
Option =2
Begin InputTables
    Name ="BuyersProposalUnitCount"
End
Begin OutputColumns
    Expression ="BuyersProposalUnitCount.SalesPersonID"
    Expression ="BuyersProposalUnitCount.UnitID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query12].[SalesPersonID]=67)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="BuyersProposalUnitCount.SalesPersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalUnitCount.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1482
    Bottom =1021
    Left =-1
    Top =-1
    Right =1450
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =61
        Top =174
        Right =205
        Bottom =318
        Top =0
        Name ="BuyersProposalUnitCount"
        Name =""
    End
End
