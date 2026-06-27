Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposalsStatus"
End
Begin OutputColumns
    Expression ="BuyersProposalsStatus.BuyerID"
    Expression ="BuyersProposalsStatus.ProposalID"
    Alias ="Units"
    Expression ="Count(BuyersProposalsStatus.UnitID)"
    Alias ="Sold"
    Expression ="Sum(BuyersProposalsStatus.UnitsSoldCount)"
    Alias ="NotSold"
    Expression ="Sum(BuyersProposalsStatus.UnitsNotSold)"
    Alias ="Isold"
    Expression ="Sum(BuyersProposalsStatus.SelfSoldCount)"
    Alias ="ProposalStatus"
    Expression ="Switch([NotSold]>0,\"Active\",[Isold]>0 And [Sold]=[Units],\"Sold\",[Sold]=[Unit"
        "s],\"Closed\")"
    Alias ="LastAction"
    Expression ="Max(Nz([LastContactDate],[FirstContactDate]))"
End
Begin Groups
    Expression ="BuyersProposalsStatus.BuyerID"
    GroupLevel =0
    Expression ="BuyersProposalsStatus.ProposalID"
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
        dbText "Name" ="ProposalStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalsStatus.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposalsStatus.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastAction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NotSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Isold"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =109
    Top =27
    Right =1610
    Bottom =1008
    Left =-1
    Top =-1
    Right =1469
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =453
        Top =117
        Right =822
        Bottom =386
        Top =0
        Name ="BuyersProposalsStatus"
        Name =""
    End
End
