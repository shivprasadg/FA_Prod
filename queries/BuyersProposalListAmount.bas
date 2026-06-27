Operation =1
Option =0
Begin InputTables
    Name ="BuyersProposal"
    Name ="Buyers"
    Name ="BuyersProposalUnits"
    Name ="Rmkt"
    Name ="BuyerProposalAutoStatus"
End
Begin OutputColumns
    Expression ="BuyersProposal.ProposalID"
    Alias ="SP"
    Expression ="UCase(Left([SalesPerson],2))"
    Alias ="Buyer"
    Expression ="Buyers.ClientShNm"
    Expression ="BuyersProposal.ProposalName"
    Alias ="Status"
    Expression ="BuyerProposalAutoStatus.ProposalStatus"
    Alias ="Last"
    Expression ="Format(Nz([LastContactDate],[FirstContactDate]),\"mm/dd/yy\")"
    Alias ="Units"
    Expression ="Count(Rmkt.UnitRef)"
    Alias ="Amount"
    Expression ="Format(Sum([TargetSale]),\"$#,###\")"
    Expression ="BuyersProposal.BuyerID"
    Expression ="BuyersProposal.SalesPersonID"
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
    LeftTable ="BuyersProposal"
    RightTable ="BuyerProposalAutoStatus"
    Expression ="BuyersProposal.ProposalID = BuyerProposalAutoStatus.ProposalID"
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
    Expression ="BuyerProposalAutoStatus.ProposalStatus"
    GroupLevel =0
    Expression ="Format(Nz([LastContactDate],[FirstContactDate]),\"mm/dd/yy\")"
    GroupLevel =0
    Expression ="BuyersProposal.BuyerID"
    GroupLevel =0
    Expression ="BuyersProposal.SalesPersonID"
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
        dbText "Name" ="BuyersProposal.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposal.BuyerID"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposal.ProposalName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyersProposal.SalesPersonID"
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
    Bottom =629
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Buyers"
        Name =""
    End
    Begin
        Left =239
        Top =170
        Right =383
        Bottom =314
        Top =0
        Name ="BuyersProposalUnits"
        Name =""
    End
    Begin
        Left =493
        Top =287
        Right =671
        Bottom =661
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =541
        Top =35
        Right =685
        Bottom =179
        Top =0
        Name ="BuyerProposalAutoStatus"
        Name =""
    End
End
