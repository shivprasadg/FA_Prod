Operation =1
Option =0
Where ="(((Rmkt.invoicepaiddate)>([forms]![acctmenu]![AsOfDate]-90)))"
Having ="(((Avg(Rmkt.SoldPrice))>0))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="Rmkt"
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Alias ="Description"
    Expression ="LookUp_UnitDesc.DescriptionMed"
    Alias ="AvgSoldPrice"
    Expression ="Avg(Rmkt.SoldPrice)"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="LookUp_UnitDesc"
    Expression ="vw_SixKeys.AssetID = LookUp_UnitDesc.AssetID"
    Flag =1
End
Begin Groups
    Expression ="LookUp_UnitDesc.DescriptionMed"
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
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="5715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgSoldPrice"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionShort"
        dbInteger "ColumnWidth" ="3315"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionMed"
        dbInteger "ColumnWidth" ="5040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =176
    Top =187
    Right =1570
    Bottom =1121
    Left =-1
    Top =-1
    Right =1362
    Bottom =566
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =273
        Top =260
        Right =417
        Bottom =404
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =592
        Top =328
        Right =736
        Bottom =645
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =45
        Top =43
        Right =212
        Bottom =442
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =260
        Top =12
        Right =502
        Bottom =235
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
