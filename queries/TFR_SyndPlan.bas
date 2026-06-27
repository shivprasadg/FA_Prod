Operation =1
Option =0
Begin InputTables
    Name ="TFR_2"
End
Begin OutputColumns
    Expression ="TFR_2.FundingStatus"
    Expression ="TFR_2.AccountName"
    Expression ="TFR_2.FAID"
    Expression ="TFR_2.Sch_Grp"
    Expression ="TFR_2.UnitDesc"
    Expression ="TFR_2.UnitCount"
    Expression ="TFR_2.FundDate"
    Expression ="TFR_2.Bank"
    Alias ="TotalOEC"
    Expression ="Sum(TFR_2.OEC)"
    Alias ="NetFee"
    Expression ="Sum(TFR_2.NetSyndFee)"
    Alias ="NetFeePct"
    Expression ="Format(Sum([NetSyndFee])/Sum([OEC]),\"Percent\")"
    Alias ="ProjFee"
    Expression ="Sum([SynFeeProj]+[StripRent]+[TermPenaltyLPAF])"
    Alias ="ProjFeePct"
    Expression ="Format([ProjFee]/Sum([OEC]),\"Percent\")"
    Expression ="TFR_2.SGrpID"
End
Begin OrderBy
    Expression ="TFR_2.FundingStatus"
    Flag =1
    Expression ="TFR_2.AccountName"
    Flag =0
    Expression ="CDbl(NumInString([Schedule])) & [UnitGroup]"
    Flag =0
    Expression ="TFR_2.FundDate"
    Flag =1
End
Begin Groups
    Expression ="TFR_2.FundingStatus"
    GroupLevel =0
    Expression ="TFR_2.AccountName"
    GroupLevel =0
    Expression ="TFR_2.FAID"
    GroupLevel =0
    Expression ="TFR_2.Sch_Grp"
    GroupLevel =0
    Expression ="TFR_2.UnitDesc"
    GroupLevel =0
    Expression ="TFR_2.UnitCount"
    GroupLevel =0
    Expression ="TFR_2.FundDate"
    GroupLevel =0
    Expression ="TFR_2.Bank"
    GroupLevel =0
    Expression ="TFR_2.SGrpID"
    GroupLevel =0
    Expression ="CDbl(NumInString([Schedule])) & [UnitGroup]"
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
        dbText "Name" ="TFR_2.AccountName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalOEC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ProjFeePct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetFeePct"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.FundingStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.Sch_Grp"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.UnitDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =144
    Top =4
    Right =1586
    Bottom =870
    Left =-1
    Top =-1
    Right =1424
    Bottom =649
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =54
        Top =35
        Right =284
        Bottom =665
        Top =0
        Name ="TFR_2"
        Name =""
    End
End
