Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposal"
End
Begin OutputColumns
    Expression ="BuyersProposal.BuyerID"
    Alias ="ProposalCount"
    Expression ="Count(BuyersProposal.ProposalID)"
End
Begin Groups
    Expression ="BuyersProposal.BuyerID"
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
        dbText "Name" ="ProposalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposal.BuyerID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1213
    Bottom =1021
    Left =-1
    Top =-1
    Right =1093
    Bottom =540
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =62
        Top =28
        Right =331
        Bottom =293
        Top =0
        Name ="BuyersProposal"
        Name =""
    End
End
