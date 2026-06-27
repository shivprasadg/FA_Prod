Operation =1
Option =0
Begin InputTables
    Name ="Client_Lease_Detail"
End
Begin OutputColumns
    Expression ="Client_Lease_Detail.Client"
    Expression ="Client_Lease_Detail.Fleet_Advantage_Agreement"
    Expression ="Client_Lease_Detail.[Assignee Contract No]"
    Expression ="Client_Lease_Detail.[Assignee Bank]"
    Expression ="Client_Lease_Detail.UnitDesc"
    Expression ="Client_Lease_Detail.Location"
    Alias ="Units"
    Expression ="Count(Client_Lease_Detail.VIN)"
    Expression ="Client_Lease_Detail.Commencement_Date"
    Expression ="Client_Lease_Detail.Final_Payment_Date"
    Alias ="Lease Rate Factor"
    Expression ="IIf(Sum([SyndRent])>0,Round(Sum([SyndRent])/Sum([Total Cost]),5),0)"
    Alias ="Estimated Rent"
    Expression ="Sum(Client_Lease_Detail.SyndRent)"
    Expression ="Client_Lease_Detail.ClientGroupId"
    Expression ="Client_Lease_Detail.SchID"
End
Begin OrderBy
    Expression ="Client_Lease_Detail.SchID"
    Flag =0
End
Begin Groups
    Expression ="Client_Lease_Detail.Client"
    GroupLevel =0
    Expression ="Client_Lease_Detail.Fleet_Advantage_Agreement"
    GroupLevel =0
    Expression ="Client_Lease_Detail.[Assignee Contract No]"
    GroupLevel =0
    Expression ="Client_Lease_Detail.[Assignee Bank]"
    GroupLevel =0
    Expression ="Client_Lease_Detail.UnitDesc"
    GroupLevel =0
    Expression ="Client_Lease_Detail.Location"
    GroupLevel =0
    Expression ="Client_Lease_Detail.Commencement_Date"
    GroupLevel =0
    Expression ="Client_Lease_Detail.Final_Payment_Date"
    GroupLevel =0
    Expression ="Client_Lease_Detail.ClientGroupId"
    GroupLevel =0
    Expression ="Client_Lease_Detail.SchID"
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
        dbText "Name" ="Units"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.Fleet_Advantage_Agreement"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.Client"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.Final_Payment_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Lease Rate Factor"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Estimated Rent"
        dbInteger "ColumnWidth" ="3780"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.Commencement_Date"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.Location"
        dbInteger "ColumnWidth" ="3780"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.[Assignee Bank]"
        dbInteger "ColumnWidth" ="3780"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.[Assignee Contract No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2790"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Client_Lease_Detail.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =2
    Top =88
    Right =1485
    Bottom =1069
    Left =-1
    Top =-1
    Right =1457
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =365
        Bottom =273
        Top =0
        Name ="Client_Lease_Detail"
        Name =""
    End
End
