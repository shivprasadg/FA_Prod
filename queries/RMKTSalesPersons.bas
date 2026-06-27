dbMemo "SQL" ="SELECT TBL_Employees.ID, TBL_Employees.[Employee Name] AS SalesPerson, TBL_Emplo"
    "yees.EmpInt AS SP, TBL_Employees.[User ID], TBL_Employees.[Access Level], CInt(N"
    "z([Events_ByType_Last30].[callsIn],0)) AS CallsIn, CInt(Nz([Events_ByType_Last30"
    "].[CallsOut],0)) AS CallsOut, CInt(Nz([Events_ByType_Last30].[EmailsIn],0)) AS E"
    "mailsIn, CInt(Nz([Events_ByType_Last30].[EmailsOut],0)) AS EmailsOut, CInt(Nz([E"
    "vents_ByType_Last30].[VisitsScheduled],0)) AS VisitsScheduled, CInt(Nz([BuyersPr"
    "oposalstatus_Last30].[Proposals],0)) AS Proposals, CInt(Nz([BuyersProposalstatus"
    "_last30].[ProposalActive],0)) AS ProposalsActive, CInt(Nz([BuyersProposalstatus_"
    "Last30].[Proposalsold],0)) AS ProposalsSold, CInt(Nz([UnitsSoldLast30],0)) AS So"
    "ld30, Format(Nz([AmtSold30],0),\"$#,###\") AS Sales30, CInt(Nz([UnitsSoldLast60]"
    ",0)) AS Sold60, Format(Nz([AmtSold60],0),\"$#,###\") AS Sales60, CInt(Nz([UnitsS"
    "oldLast90],0)) AS Sold90, Format(Nz([AmtSold90],0),\"$#,###\") AS Sales90, Event"
    "s_ByType_Last30.LastCall, Events_ByType_Last30.LastEmail, Events_ByType_Last30.L"
    "astVisit, Events_ByType_Last30.NextVisit, Events_ByType_Last30.EmailsPending, Ev"
    "ents_ByType_Last30.CallsPending\015\012FROM ((TBL_Employees LEFT JOIN BuyersProp"
    "osalstatus_Last30 ON TBL_Employees.ID = BuyersProposalstatus_Last30.SalesPersonI"
    "D) LEFT JOIN Events_ByType_Last30 ON TBL_Employees.ID = Events_ByType_Last30.Sal"
    "esPersonID) LEFT JOIN RMKTUnitsSold30_60_90 ON TBL_Employees.EmpInt = RMKTUnitsS"
    "old30_60_90.RmkSm\015\012WHERE (((TBL_Employees.AccessGroup) Like \"*RMKT*\" Or "
    "(TBL_Employees.AccessGroup) Like \"*ALL*\"))\015\012ORDER BY TBL_Employees.[Empl"
    "oyee Name];\015\012"
dbMemo "Connect" =""
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="TBL_Employees.ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="TBL_Employees.[User ID]"
        dbInteger "ColumnOrder" ="8"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CallsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VisitsScheduled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales90"
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
End
