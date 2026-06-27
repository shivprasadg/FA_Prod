Operation =1
Option =0
Begin InputTables
    Name ="BDVOP_Excel1"
    Name ="BDVOP_Excel0"
End
Begin OutputColumns
    Alias ="SentYear"
    Expression ="(Nz([BDVOP_Excel1].[SentYear],[BDVOP_Excel1].[DealYear]))"
    Expression ="BDVOP_Excel1.Client"
    Expression ="BDVOP_Excel1.FAID"
    Expression ="BDVOP_Excel1.DealBEDEX"
    Expression ="BDVOP_Excel1.DealYear"
    Alias ="DealUnitQty"
    Expression ="CInt(Nz([BDVOP_Excel1].[DealUnitQty],0))"
    Expression ="BDVOP_Excel1.DealGrandTotal"
    Alias ="Units"
    Expression ="CInt(Nz([UnitsInGroup],0))"
    Alias ="OEC"
    Expression ="CCur(Nz([OrigEquipCost],0))"
    Alias ="UnitDiff"
    Expression ="(CInt(Nz([UnitsInGroup],0))-CInt(Nz([DealUnitQty],0)))"
    Alias ="OECDiff"
    Expression ="[OEC]-[DealGrandTotal]"
    Expression ="BDVOP_Excel1.OutcomeDate"
End
Begin Joins
    LeftTable ="BDVOP_Excel1"
    RightTable ="BDVOP_Excel0"
    Expression ="BDVOP_Excel1.FAID = BDVOP_Excel0.LFAID"
    Flag =2
End
Begin OrderBy
    Expression ="(Nz([BDVOP_Excel1].[SentYear],[BDVOP_Excel1].[DealYear]))"
    Flag =1
    Expression ="BDVOP_Excel1.Client"
    Flag =0
End
Begin Groups
    Expression ="(Nz([BDVOP_Excel1].[SentYear],[BDVOP_Excel1].[DealYear]))"
    GroupLevel =0
    Expression ="BDVOP_Excel1.Client"
    GroupLevel =0
    Expression ="BDVOP_Excel1.FAID"
    GroupLevel =0
    Expression ="BDVOP_Excel1.DealBEDEX"
    GroupLevel =0
    Expression ="BDVOP_Excel1.DealYear"
    GroupLevel =0
    Expression ="CInt(Nz([BDVOP_Excel1].[DealUnitQty],0))"
    GroupLevel =0
    Expression ="BDVOP_Excel1.DealGrandTotal"
    GroupLevel =0
    Expression ="CInt(Nz([UnitsInGroup],0))"
    GroupLevel =0
    Expression ="CCur(Nz([OrigEquipCost],0))"
    GroupLevel =0
    Expression ="(CInt(Nz([UnitsInGroup],0))-CInt(Nz([DealUnitQty],0)))"
    GroupLevel =0
    Expression ="BDVOP_Excel1.OutcomeDate"
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
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SentYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="DealUnitQty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="UnitDiff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OECDiff"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.DealYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.DealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.FAID"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.Client"
        dbInteger "ColumnWidth" ="3675"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.OutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDVOP_Excel1.DealBEDEX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-437
    Top =240
    Right =1029
    Bottom =886
    Left =-1
    Top =-1
    Right =1442
    Bottom =258
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =30
        Top =25
        Right =194
        Bottom =227
        Top =0
        Name ="BDVOP_Excel1"
        Name =""
    End
    Begin
        Left =331
        Top =73
        Right =475
        Bottom =217
        Top =0
        Name ="BDVOP_Excel0"
        Name =""
    End
End
