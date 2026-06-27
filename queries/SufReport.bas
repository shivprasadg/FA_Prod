Operation =1
Option =0
Where ="(((Rmkt.[SUF Sent]) Is Null))"
Begin InputTables
    Name ="vw_Sch_Grp_Units_Remarketing"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="vw_Sch_Grp_Units_Remarketing.clientID"
    Expression ="vw_Sch_Grp_Units_Remarketing.CompanyName"
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID"
    Expression ="vw_Sch_Grp_Units_Remarketing.VIN"
    Expression ="Rmkt.BuyOutCost"
    Expression ="Rmkt.BuyOutDate"
    Expression ="Rmkt.[SUF Sent]"
    Expression ="Rmkt.[SUF Returned]"
    Expression ="vw_Sch_Grp_Units_Remarketing.Status"
End
Begin Joins
    LeftTable ="vw_Sch_Grp_Units_Remarketing"
    RightTable ="Rmkt"
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID = Rmkt.UnitRef"
    Flag =1
End
Begin OrderBy
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID"
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
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.CompanyName"
        dbInteger "ColumnWidth" ="4260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.AssetID"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.VIN"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =51
    Top =2
    Right =1311
    Bottom =771
    Left =-1
    Top =-1
    Right =1228
    Bottom =452
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =14
        Top =15
        Right =226
        Bottom =473
        Top =0
        Name ="vw_Sch_Grp_Units_Remarketing"
        Name =""
    End
    Begin
        Left =296
        Top =112
        Right =440
        Bottom =440
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
