Operation =1
Option =0
Begin InputTables
    Name ="RMKTSalesPersons_Last30"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="Count(RMKTSalesPersons_Last30.ID)"
    Alias ="CallsIn"
    Expression ="Sum(RMKTSalesPersons_Last30.CallsIn)"
    Alias ="CallsOut"
    Expression ="Sum(RMKTSalesPersons_Last30.CallsOut)"
    Alias ="EmailsIn"
    Expression ="Sum(RMKTSalesPersons_Last30.EmailsIn)"
    Alias ="EmailsOut"
    Expression ="Sum(RMKTSalesPersons_Last30.EmailsOut)"
    Alias ="VisitsScheduled"
    Expression ="Sum(RMKTSalesPersons_Last30.VisitsScheduled)"
    Alias ="Proposals"
    Expression ="Sum(RMKTSalesPersons_Last30.Proposals)"
    Alias ="ProposalsActive"
    Expression ="Sum(RMKTSalesPersons_Last30.ProposalsActive)"
    Alias ="ProposalsSold"
    Expression ="Sum(RMKTSalesPersons_Last30.ProposalsSold)"
    Alias ="Sold30"
    Expression ="Sum(RMKTSalesPersons_Last30.Sold30)"
    Alias ="Sales30"
    Expression ="Format(Sum([RMKTSalesPersons_Last30].[Sales30]),\"$#,##0\")"
    Alias ="LastCall"
    Expression ="Max(RMKTSalesPersons_Last30.LastCall)"
    Alias ="LastEmail"
    Expression ="Max(RMKTSalesPersons_Last30.LastEmail)"
    Alias ="LastVisit"
    Expression ="Max(RMKTSalesPersons_Last30.LastVisit)"
    Alias ="NextVisit"
    Expression ="Max(RMKTSalesPersons_Last30.NextVisit)"
    Alias ="EmailsPending"
    Expression ="Sum(RMKTSalesPersons_Last30.EmailsPending)"
    Alias ="CallsPending"
    Expression ="Sum(RMKTSalesPersons_Last30.CallsPending)"
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
        dbText "Name" ="CallsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VisitsScheduled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProposalsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProposalsActive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSales30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsOut"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.SalesPerson"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.LastCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.ProposalsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSales60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.SP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.VisitsScheduled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.CallsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfProposalsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.NextVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.LastEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.CallsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NextVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSales90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.ProposalsActive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.CallsOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sales90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sales60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.LastVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.EmailsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfSold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLastCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCallsOut"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.EmailsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfProposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sales30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCallsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsPending"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Sold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.EmailsOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLastEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCallsPending"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLastVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfEmailsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfEmailsOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons_Last30.Proposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfNextVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfEmailsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =102
    Top =25
    Right =1976
    Bottom =1006
    Left =-1
    Top =-1
    Right =1842
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =503
        Top =82
        Right =916
        Bottom =405
        Top =0
        Name ="RMKTSalesPersons_Last30"
        Name =""
    End
End
