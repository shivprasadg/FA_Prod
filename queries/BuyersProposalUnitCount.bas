Operation =1
Option =2
Begin InputTables
    Name ="BuyersProposalUnits"
    Name ="BuyersProposal"
End
Begin OutputColumns
    Expression ="BuyersProposal.SalesPersonID"
    Expression ="BuyersProposalUnits.ProposalID"
    Expression ="BuyersProposalUnits.UnitID"
End
Begin Joins
    LeftTable ="BuyersProposalUnits"
    RightTable ="BuyersProposal"
    Expression ="BuyersProposalUnits.ProposalID = BuyersProposal.ProposalID"
    Flag =1
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
        dbText "Name" ="BuyersProposal.SalesPersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalUnits.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalUnits.UnitID"
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
    Bottom =472
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =73
        Top =34
        Right =217
        Bottom =178
        Top =0
        Name ="BuyersProposalUnits"
        Name =""
    End
    Begin
        Left =236
        Top =56
        Right =477
        Bottom =402
        Top =0
        Name ="BuyersProposal"
        Name =""
    End
End
