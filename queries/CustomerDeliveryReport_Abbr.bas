Operation =1
Option =0
Begin InputTables
    Name ="vw_CustDelvReportSummed"
End
Begin OutputColumns
    Expression ="vw_CustDelvReportSummed.Client"
    Alias ="Sch|Grp"
    Expression ="vw_CustDelvReportSummed.schgrp"
    Alias ="#"
    Expression ="IIf([units]<>[UnitsInGroup],CStr([Units]) & \" (\" & [UnitsInGroup] & \")\",[Uni"
        "tsInGroup])"
    Alias ="Unit Description"
    Expression ="vw_CustDelvReportSummed.UnitDescAbbr"
    Expression ="vw_CustDelvReportSummed.Vendor"
    Alias ="Build"
    Expression ="Format([vw_CustDelvReportSummed].[Build],\"m/d/yy\")"
    Expression ="vw_CustDelvReportSummed.Built"
    Expression ="vw_CustDelvReportSummed.[Est Delv Dates]"
    Expression ="vw_CustDelvReportSummed.Delv"
    Expression ="vw_CustDelvReportSummed.MSO"
    Expression ="vw_CustDelvReportSummed.INV"
    Expression ="vw_CustDelvReportSummed.Assignee"
    Alias ="PO"
    Expression ="IIf([CustPO1]=True,\"C\",IIf(IsDate([PO Date]),\"Y\",\"\"))"
    Alias ="IntDelv"
    Expression ="IIf(Nz([Interim bank],\"\")<>\"\" And [FundDate]>Date(),1,0)"
    Alias ="IntBank"
    Expression ="vw_CustDelvReportSummed.[Interim Bank]"
End
Begin OrderBy
    Expression ="IIf(Nz([Interim bank],\"\")<>\"\" And [FundDate]>Date(),1,0)"
    Flag =0
    Expression ="vw_CustDelvReportSummed.Client"
    Flag =0
    Expression ="vw_CustDelvReportSummed.schgrp"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Sch|Grp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit Description"
        dbInteger "ColumnWidth" ="4065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Build"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.Delv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.Vendor"
        dbInteger "ColumnWidth" ="3915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.Built"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.[Est Delv Dates]"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.MSO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportSummed.INV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IntDelv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IntBank"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-353
    Top =18
    Right =1080
    Bottom =753
    Left =-1
    Top =-1
    Right =1409
    Bottom =418
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =239
        Top =41
        Right =502
        Bottom =363
        Top =0
        Name ="vw_CustDelvReportSummed"
        Name =""
    End
End
