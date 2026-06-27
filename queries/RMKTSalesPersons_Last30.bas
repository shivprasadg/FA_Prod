Operation =1
Option =0
Where ="(((TBL_Employees.AccessGroup) Like \"*RMKT*\" Or (TBL_Employees.AccessGroup) Lik"
    "e \"*ALL*\"))"
Begin InputTables
    Name ="TBL_Employees"
    Name ="BuyersProposalStatus_Last30"
    Name ="Events_ByType_Last30"
    Name ="RMKTUnitsSold30_60_90"
End
Begin OutputColumns
    Expression ="TBL_Employees.ID"
    Alias ="SalesPerson"
    Expression ="TBL_Employees.[Employee Name]"
    Alias ="SP"
    Expression ="TBL_Employees.EmpInt"
    Expression ="TBL_Employees.[User ID]"
    Expression ="TBL_Employees.[Access Level]"
    Alias ="CallsIn"
    Expression ="CInt(Nz([Events_ByType_Last30].[callsIn],0))"
    Alias ="CallsOut"
    Expression ="CInt(Nz([Events_ByType_Last30].[CallsOut],0))"
    Alias ="EmailsIn"
    Expression ="CInt(Nz([Events_ByType_Last30].[EmailsIn],0))"
    Alias ="EmailsOut"
    Expression ="CInt(Nz([Events_ByType_Last30].[EmailsOut],0))"
    Alias ="VisitsScheduled"
    Expression ="CInt(Nz([Events_ByType_Last30].[VisitsScheduled],0))"
    Alias ="Proposals"
    Expression ="CInt(Nz([BuyersProposalStatus_Last30].[Proposals],0))"
    Alias ="ProposalsActive"
    Expression ="CInt(Nz([BuyersProposalStatus_Last30].[ProposalActive],0))"
    Alias ="ProposalsSold"
    Expression ="CInt(Nz([BuyersProposalStatus_Last30].[Proposalsold],0))"
    Alias ="Sold30"
    Expression ="CInt(Nz([UnitsSoldLast30],0))"
    Alias ="Sales30"
    Expression ="Format(Nz([AmtSold30],0),\"$#,##0\")"
    Alias ="Sold60"
    Expression ="CInt(Nz([UnitsSoldLast60],0))"
    Alias ="Sales60"
    Expression ="Format(Nz([AmtSold60],0),\"$#,##0\")"
    Alias ="Sold90"
    Expression ="CInt(Nz([UnitsSoldLast90],0))"
    Alias ="Sales90"
    Expression ="Format(Nz([AmtSold90],0),\"$#,##0\")"
    Expression ="Events_ByType_Last30.LastCall"
    Expression ="Events_ByType_Last30.LastEmail"
    Expression ="Events_ByType_Last30.LastVisit"
    Expression ="Events_ByType_Last30.NextVisit"
    Expression ="Events_ByType_Last30.EmailsPending"
    Expression ="Events_ByType_Last30.CallsPending"
End
Begin Joins
    LeftTable ="TBL_Employees"
    RightTable ="BuyersProposalStatus_Last30"
    Expression ="TBL_Employees.ID = BuyersProposalStatus_Last30.SalesPersonID"
    Flag =2
    LeftTable ="TBL_Employees"
    RightTable ="Events_ByType_Last30"
    Expression ="TBL_Employees.ID = Events_ByType_Last30.SalesPersonID"
    Flag =2
    LeftTable ="TBL_Employees"
    RightTable ="RMKTUnitsSold30_60_90"
    Expression ="TBL_Employees.EmpInt = RMKTUnitsSold30_60_90.RmkSm"
    Flag =2
End
Begin OrderBy
    Expression ="TBL_Employees.[Employee Name]"
    Flag =0
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
        dbText "Name" ="TBL_Employees.[Access Level]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBL_Employees.ID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBL_Employees.[User ID]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesPerson"
        dbInteger "ColumnWidth" ="1950"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SP"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsOut"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsIn"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsOut"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VisitsScheduled"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsIn"
        dbInteger "ColumnOrder" ="6"
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
        dbText "Name" ="Sales30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="proposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="proposalsActive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="proposalsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerposalsActive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.EmailsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perposals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.LastCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.EmailsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.CallsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.CallsPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.LastEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerposalsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.NextVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.NextVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.LastCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.LastVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsEvent_ByType_Last30.LastVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events_ByType_Last30.LastEmail"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1225
    Top =40
    Right =1128
    Bottom =677
    Left =-1
    Top =-1
    Right =2321
    Bottom =114
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =44
        Top =-13
        Right =383
        Bottom =321
        Top =0
        Name ="TBL_Employees"
        Name =""
    End
    Begin
        Left =431
        Top =12
        Right =575
        Bottom =156
        Top =0
        Name ="BuyersProposalStatus_Last30"
        Name =""
    End
    Begin
        Left =623
        Top =12
        Right =767
        Bottom =156
        Top =0
        Name ="Events_ByType_Last30"
        Name =""
    End
    Begin
        Left =861
        Top =43
        Right =1082
        Bottom =251
        Top =0
        Name ="RMKTUnitsSold30_60_90"
        Name =""
    End
End
