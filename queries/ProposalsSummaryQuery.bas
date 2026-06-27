Operation =1
Option =0
Where ="(((Deals.transTypeRef)<>1))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Deals.transTypeRef"
    Expression ="Deals.pipelineOrder"
    Expression ="Deals.FAID"
    Expression ="Deals.dealID"
    Expression ="Deals.clientRef"
    Alias ="CompanyAndDBA"
    Expression ="[clientcompanyName] & IIf([clientDBA]<>\"\",\" d/b/a \" & [clientDBA],\"\")"
    Expression ="Clients.clientCompanyName"
    Expression ="Deals.dealEnteredDate"
    Expression ="Deals.dealEstDelDate"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeDate"
    Alias ="Edit"
    Expression ="\"View/Edit\""
    Expression ="Deals.dealShortDesc"
    Expression ="Deals.dealRFPDate"
    Expression ="Deals.dealLPAFDue"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.ProposalTimeline"
    Expression ="Deals.dealOutcomeReason"
    Expression ="Deals.dealSchedNum"
    Expression ="Deals.ExcludeFromKPI"
    Expression ="Deals.DealLOEx"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
End
Begin OrderBy
    Expression ="Deals.FAID"
    Flag =1
    Expression ="Deals.dealID"
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
        dbText "Name" ="Deals.dealID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealEnteredDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.proposalTimeline"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.pipelineOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompanyAndDBA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFDue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.ExcludeFromKPI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.DealLOEx"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1460
    Bottom =1029
    Left =-1
    Top =-1
    Right =1442
    Bottom =352
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =25
        Top =9
        Right =221
        Bottom =229
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =270
        Top =1
        Right =529
        Bottom =290
        Top =0
        Name ="Deals"
        Name =""
    End
End
