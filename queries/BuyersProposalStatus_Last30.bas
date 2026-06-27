Operation =1
Option =0
Where ="(((BuyersProposal.FirstContactDate)>=DateAdd(\"m\",-1,Date())))"
Begin InputTables
    Name ="BuyersProposal"
End
Begin OutputColumns
    Expression ="BuyersProposal.SalesPersonID"
    Expression ="BuyersProposal.SalesPerson"
    Alias ="Proposals"
    Expression ="Count(BuyersProposal.ProposalID)"
    Alias ="ProposalActive"
    Expression ="Sum(IIf([ProposalStatus] Not Like '*sold*' And [ProposalStatus]<>'Closed',1,0))"
    Alias ="ProposalSold"
    Expression ="Sum(IIf([ProposalStatus] Like '*sold*',1,0))"
End
Begin Groups
    Expression ="BuyersProposal.SalesPersonID"
    GroupLevel =0
    Expression ="BuyersProposal.SalesPerson"
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
        dbText "Name" ="proposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="proposalsold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyersproposal.SalesPersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyersproposal.SalesPerson"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProposalActive"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1230
    Bottom =1021
    Left =-1
    Top =-1
    Right =1154
    Bottom =559
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="BuyersProposal"
        Name =""
    End
End
