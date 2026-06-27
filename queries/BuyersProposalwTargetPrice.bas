Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposal"
    Name ="Buyers"
    Name ="Rmkt"
    Name ="BuyersProposalUnits"
End
Begin OutputColumns
    Expression ="BuyersProposal.ProposalID"
    Alias ="SP"
    Expression ="UCase(Left([SalesPerson],2))"
    Alias ="Buyer"
    Expression ="Buyers.ClientShNm"
    Expression ="BuyersProposal.ProposalName"
    Alias ="Status"
    Expression ="BuyersProposal.ProposalStatus"
    Alias ="Last Contact"
    Expression ="Nz([LastContactDate],[FirstContactDate])"
    Expression ="BuyersProposal.BuyerID"
    Alias ="Units"
    Expression ="Count(BuyersProposalUnits.UnitID)"
    Alias ="Target"
    Expression ="Sum(Rmkt.TargetSale)"
End
Begin Joins
    LeftTable ="BuyersProposal"
    RightTable ="Buyers"
    Expression ="BuyersProposal.BuyerID = Buyers.clientID"
    Flag =1
    LeftTable ="BuyersProposal"
    RightTable ="BuyersProposalUnits"
    Expression ="BuyersProposal.ProposalID = BuyersProposalUnits.ProposalID"
    Flag =1
    LeftTable ="BuyersProposalUnits"
    RightTable ="Rmkt"
    Expression ="BuyersProposalUnits.UnitID = Rmkt.UnitRef"
    Flag =1
End
Begin Groups
    Expression ="BuyersProposal.ProposalID"
    GroupLevel =0
    Expression ="UCase(Left([SalesPerson],2))"
    GroupLevel =0
    Expression ="Buyers.ClientShNm"
    GroupLevel =0
    Expression ="BuyersProposal.ProposalName"
    GroupLevel =0
    Expression ="BuyersProposal.ProposalStatus"
    GroupLevel =0
    Expression ="Nz([LastContactDate],[FirstContactDate])"
    GroupLevel =0
    Expression ="BuyersProposal.BuyerID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="BuyersProposal.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposal.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SP"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BuyersProposal.ProposalName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Target"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1490
    Bottom =1021
    Left =-1
    Top =-1
    Right =1458
    Bottom =306
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
    Begin
        Left =49
        Top =165
        Right =193
        Bottom =309
        Top =0
        Name ="Buyers"
        Name =""
    End
    Begin
        Left =413
        Top =13
        Right =612
        Bottom =317
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =238
        Top =12
        Right =382
        Bottom =156
        Top =0
        Name ="BuyersProposalUnits"
        Name =""
    End
End
