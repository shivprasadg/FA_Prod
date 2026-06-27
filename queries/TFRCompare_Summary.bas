Operation =1
Option =0
Begin InputTables
    Name ="TFRCompare"
End
Begin OutputColumns
    Alias ="FundStatus"
    Expression ="IIf([Assigned_Indicator]=\"F\",\"Syndicated\",IIf([Assigned_Indicator]=\"P\",\"F"
        "orecast\",\"Portfolio\"))"
    Expression ="TFRCompare.AccountName"
    Alias ="UnitCount"
    Expression ="Sum(TFRCompare.UnitCount)"
    Alias ="OEC"
    Expression ="Sum(TFRCompare.GroupOEC)"
    Alias ="SyndFee"
    Expression ="Sum(TFRCompare.SyndFee)"
    Alias ="NetSyndFee"
    Expression ="Sum(TFRCompare.NetSyndFee)"
    Alias ="NetSyndProj"
    Expression ="Sum(TFRCompare.NetSyndProj)"
    Alias ="TargetBank"
    Expression ="TFRCompare.Bank"
    Alias ="PDRent"
    Expression ="Sum(TFRCompare.PDRent)"
    Alias ="PDRentProj"
    Expression ="Sum(TFRCompare.PDRentProj)"
    Alias ="UTA"
    Expression ="Sum(TFRCompare.UTA)"
    Alias ="UTAProj"
    Expression ="Sum(TFRCompare.UTAProj)"
    Alias ="StripRent"
    Expression ="Sum(TFRCompare.StripRent)"
    Alias ="TotalFee"
    Expression ="Sum(TFRCompare.TotalFee)"
    Alias ="TotalFeeProj"
    Expression ="Sum(TFRCompare.TotalFeeProj)"
    Expression ="TFRCompare.clientID"
    Expression ="TFRCompare.Assigned_Indicator"
    Alias ="TotPDRent"
    Expression ="Sum(TFRCompare.PDRent)"
    Alias ="AvgPDDays"
    Expression ="Avg(TFRCompare.PDDays)"
    Alias ="RentIncome"
    Expression ="Sum(TFRCompare.PmtAmt)"
    Alias ="SyndFeePlus"
    Expression ="Sum(TFRCompare.SyndFeePlus)"
End
Begin OrderBy
    Expression ="IIf([Assigned_Indicator]=\"F\",\"Syndicated\",IIf([Assigned_Indicator]=\"P\",\"F"
        "orecast\",\"Portfolio\"))"
    Flag =1
    Expression ="TFRCompare.AccountName"
    Flag =0
End
Begin Groups
    Expression ="IIf([Assigned_Indicator]=\"F\",\"Syndicated\",IIf([Assigned_Indicator]=\"P\",\"F"
        "orecast\",\"Portfolio\"))"
    GroupLevel =0
    Expression ="TFRCompare.AccountName"
    GroupLevel =0
    Expression ="TFRCompare.Bank"
    GroupLevel =0
    Expression ="TFRCompare.clientID"
    GroupLevel =0
    Expression ="TFRCompare.Assigned_Indicator"
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
        dbText "Name" ="FundStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalFeeProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRentProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTAProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetSyndProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetSyndFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccountName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.PmtAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentIncome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.PDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.PDDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotPDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfPDDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgPDDays"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPmtAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.AccountName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.SyndFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSyndFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.StripRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StripRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSyndFeePlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TFRCompare.SyndFeePlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndFeePlus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-35
    Top =59
    Right =1167
    Bottom =892
    Left =-1
    Top =-1
    Right =1178
    Bottom =386
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =302
        Bottom =234
        Top =0
        Name ="TFRCompare"
        Name =""
    End
End
