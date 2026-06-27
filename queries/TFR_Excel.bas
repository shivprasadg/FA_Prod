Operation =1
Option =0
Begin InputTables
    Name ="TFR_2"
End
Begin OutputColumns
    Expression ="TFR_2.AccountName"
    Expression ="TFR_2.Sch_Grp"
    Expression ="TFR_2.UnitDesc"
    Expression ="TFR_2.UnitCount"
    Expression ="TFR_2.FundDate"
    Alias ="FundMthYear"
    Expression ="Format([TFR_2].[FundDate],\"mmm yyyy\")"
    Expression ="TFR_2.OEC"
    Expression ="TFR_2.OECPerUnit"
    Expression ="TFR_2.Bank"
    Expression ="TFR_2.SyndFee"
    Expression ="TFR_2.SynFeePct"
    Expression ="TFR_2.TermPenalty"
    Expression ="TFR_2.TermPenaltyPct"
    Expression ="TFR_2.NetSyndFee"
    Expression ="TFR_2.NetSyndPct"
    Expression ="TFR_2.PDRent"
    Expression ="TFR_2.PDRentPct"
    Expression ="TFR_2.PDDaysAvg"
    Expression ="TFR_2.UTA"
    Expression ="TFR_2.UTApct"
    Expression ="TFR_2.NPV"
    Expression ="TFR_2.NPVPct"
    Expression ="TFR_2.NetFee"
    Expression ="TFR_2.NetFeePct"
    Expression ="TFR_2.PmtAmt"
    Expression ="TFR_2.FMSS"
    Expression ="TFR_2.PmtAmtFMSS"
    Expression ="TFR_2.RVPerUnit"
    Expression ="TFR_2.RVPct"
    Expression ="TFR_2.LeaseTermPrimary"
    Expression ="TFR_2.MilesAllowed"
    Expression ="TFR_2.Assigned_Indicator"
    Expression ="TFR_2.FundDate"
    Expression ="TFR_2.SWAP"
    Expression ="TFR_2.Spread"
    Expression ="TFR_2.Yield"
    Expression ="TFR_2.Bonus"
End
Begin OrderBy
    Expression ="TFR_2.AccountName"
    Flag =0
    Expression ="CDbl(NumInString([Schedule])) & [UnitGroup]"
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
dbInteger "RowHeight" ="330"
Begin
    Begin
        dbText "Name" ="TFR_2.UnitDesc"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.NetSyndPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.SyndFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.TermPenalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.TermPenaltyPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.NetSyndFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.UTA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.NetFeePct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.UTApct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.NetFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TFR_2.NPV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.MilesAllowed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.FMSS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.PDRentPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.PDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PmtAmt"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FundMthYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.AccountName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.RVPerUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.RVPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.NPVPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.SynFeePct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.PDDaysAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.Bonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.Sch_Grp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.OECPerUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.Yield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.PmtAmt"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFR_2.PmtAmtFMSS"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =135
    Top =-5
    Right =1607
    Bottom =989
    Left =-1
    Top =-1
    Right =1454
    Bottom =494
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =86
        Top =-4
        Right =459
        Bottom =459
        Top =0
        Name ="TFR_2"
        Name =""
    End
End
